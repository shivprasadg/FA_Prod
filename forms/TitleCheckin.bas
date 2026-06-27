Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15120
    DatasheetFontHeight =11
    ItemSuffix =88
    Left =5670
    Top =5565
    Right =20790
    Bottom =17835
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf72b63b67453e440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =645
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =480
                    Width =15120
                    Height =165
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =480
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =645
                    TabIndex =2
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =247
                    Left =15
                    Width =3375
                    Height =585
                    FontSize =24
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label8"
                    Caption ="Title Log Entry"
                    LayoutCachedLeft =15
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =585
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =12898
                    Top =150
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12898
                    LayoutCachedTop =150
                    LayoutCachedWidth =13453
                    LayoutCachedHeight =378
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =13455
                    Top =150
                    Width =750
                    Height =228
                    FontSize =8
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13455
                    LayoutCachedTop =150
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =378
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11640
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10080
                    Top =510
                    Width =780
                    Height =315
                    FontSize =10
                    TabIndex =2
                    BackColor =15073023
                    Name ="seekst"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY lookup_States.State"
                        "Abbrev;"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =10080
                    LayoutCachedTop =510
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8580
                            Top =510
                            Width =1440
                            Height =315
                            FontSize =10
                            Name ="Label19"
                            Caption ="Titled State"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =8580
                            LayoutCachedTop =510
                            LayoutCachedWidth =10020
                            LayoutCachedHeight =825
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10920
                    Top =885
                    Width =3060
                    FontSize =10
                    TabIndex =5
                    Name ="LocVINBtn"
                    Caption ="Locate VIN"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =10920
                    LayoutCachedTop =885
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1245
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8872517
                    BorderColor =8210719
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12240
                    Top =510
                    Width =1740
                    Height =315
                    FontSize =10
                    TabIndex =3
                    BackColor =15073023
                    BorderColor =-2147483615
                    Name ="seektr"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    ColumnWidths ="1440"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =12240
                    LayoutCachedTop =510
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10920
                            Top =510
                            Width =1260
                            Height =315
                            FontSize =10
                            Name ="Titled to Trust_Label"
                            Caption ="Titled to Trust"
                            FontName ="Segoe UI"
                            EventProcPrefix ="Titled_to_Trust_Label"
                            GroupTable =1
                            LayoutCachedLeft =10920
                            LayoutCachedTop =510
                            LayoutCachedWidth =12180
                            LayoutCachedHeight =825
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =9120
                    Top =4635
                    Width =5790
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="SchoutBtn"
                    Caption ="Check out Schedule to Trust"
                    FontName ="Segoe UI"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="TitletoTrustq"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"SchoutBtn\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">TitletoTrustq<"
                        End
                        Begin
                            Comment ="_AXL:/Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =9120
                    LayoutCachedTop =4635
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =4935
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverColor =8965045
                    PressedThemeColorIndex =4
                    PressedShade =50.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    Overlaps =1
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =6480
                    Top =1335
                    Width =8460
                    Height =3270
                    TabIndex =8
                    Name ="TitleCheckInSub"
                    SourceObject ="Form.TitleCheckInSub"

                    LayoutCachedLeft =6480
                    LayoutCachedTop =1335
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =4605
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =10
                    Top =60
                    Width =5775
                    Height =4440
                    FontSize =9
                    TabIndex =6
                    Name ="TitleLogList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitTitleHistoryLastEntry.UnitRef AS UnitID, UnitTitleHistoryLastEntry.Un"
                        "itTitleLogDate AS LogDate, IIf([EntryTypeId]=1,\"In\",\"Out\") AS [Action], Unit"
                        "s.UnitVIN AS VIN, Units.UnitYr AS [Year], Nz([TitleSt],[TitledLocationState]) AS"
                        " State, UnitTitleHistoryLastEntry.EntryReason AS Reason, UnitTitleHistoryLastEnt"
                        "ry.TitleNumber, UnitTitleHistoryLastEntry.ExpectedReturn, UnitTitleHistoryLastEn"
                        "try.SchID FROM UnitTitleHistoryLastEntry LEFT JOIN Units ON UnitTitleHistoryLast"
                        "Entry.UnitRef = Units.UnitID ORDER BY UnitTitleHistoryLastEntry.UnitTitleLogDate"
                        " DESC;"
                    ColumnWidths ="0;1080;720;2304;576;576;0;0;0;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedTop =60
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =4500
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7575
                    Top =510
                    Width =945
                    Height =315
                    FontSize =10
                    TabIndex =1
                    BackColor =15073023
                    BorderColor =-2147483615
                    Name ="seekMY"
                    FontName ="Segoe UI"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =7575
                    LayoutCachedTop =510
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =825
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
                            Left =6540
                            Top =510
                            Width =975
                            Height =315
                            FontSize =10
                            Name ="Label3"
                            Caption ="Enter Year"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =6540
                            LayoutCachedTop =510
                            LayoutCachedWidth =7515
                            LayoutCachedHeight =825
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =50
                    IMESentenceMode =3
                    Left =7575
                    Top =60
                    Width =6405
                    Height =420
                    FontSize =16
                    BackColor =15073023
                    BorderColor =-2147483615
                    Name ="SeekVin"
                    FontName ="OCR A Extended"
                    OnChange ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =7575
                    LayoutCachedTop =60
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =480
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6540
                            Top =60
                            Width =975
                            Height =420
                            FontSize =10
                            Name ="Label1"
                            Caption ="Enter VIN"
                            FontName ="Segoe UI"
                            GroupTable =1
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =6540
                            LayoutCachedTop =60
                            LayoutCachedWidth =7515
                            LayoutCachedHeight =480
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14160
                    Top =75
                    Width =660
                    Height =480
                    FontSize =16
                    FontWeight =700
                    TabIndex =7
                    Name ="Direction"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =14160
                    LayoutCachedTop =75
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =555
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7575
                    Top =885
                    Width =3285
                    Height =360
                    FontSize =10
                    TabIndex =4
                    BackColor =15073023
                    BorderColor =-2147483615
                    Name ="SeekReason"
                    FontName ="Segoe UI"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =7575
                    LayoutCachedTop =885
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1245
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6540
                            Top =885
                            Width =975
                            Height =360
                            FontSize =10
                            Name ="Label68"
                            Caption ="Reason "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =6540
                            LayoutCachedTop =885
                            LayoutCachedWidth =7515
                            LayoutCachedHeight =1245
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =6
                    Top =4995
                    Width =5775
                    Height =6405
                    FontSize =9
                    TabIndex =10
                    BoundColumn =4
                    Name ="TitleSchList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitTitleLog.UnitTitleLogDate AS LogDate, vw_SixKeys.ClientShortName AS C"
                        "lient, [Schedule] & \" | \" & [UnitGroup] AS SchGrp, Count(UnitTitleLog.UnitRef)"
                        " AS TitleEvents, vw_SixKeys.SchID, vw_SixKeys.SGrpID FROM UnitTitleLog INNER JOI"
                        "N vw_SixKeys ON UnitTitleLog.UnitRef = vw_SixKeys.UnitID GROUP BY UnitTitleLog.U"
                        "nitTitleLogDate, vw_SixKeys.ClientShortName, [Schedule] & \" | \" & [UnitGroup],"
                        " vw_SixKeys.SchID, vw_SixKeys.SGrpID HAVING (((UnitTitleLog.UnitTitleLogDate)>Da"
                        "te()-3)) ORDER BY UnitTitleLog.UnitTitleLogDate DESC;"
                    ColumnWidths ="1152;2160;1440;720;0;0"

                    LayoutCachedTop =4995
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =11400
                End
                Begin Label
                    OverlapFlags =93
                    Top =4620
                    Width =3375
                    Height =375
                    FontSize =16
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label80"
                    Caption ="Recent Log Entries"
                    LayoutCachedTop =4620
                    LayoutCachedWidth =3375
                    LayoutCachedHeight =4995
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12660
                    Top =7058
                    Width =1860
                    Height =278
                    FontSize =9
                    TabIndex =13
                    Name ="UpdateAllUnits"
                    Caption ="Update All Units"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ff0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b3b3b3c300000000b3b3b318b3b3b372b3b3b3d2b3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ff000000000000000000000000000000003255d696 ,
                        0x3255d6f03255d6a20000000000000000b17d4af9b17d4a78b17d4a1800000000 ,
                        0x000000000000000000000000000000000000000000000000000000003255d6cc ,
                        0x3255d6ff3255d6e1b3b3b3ffb3b3b378b17d4a81b17d4affb17d4af300000000 ,
                        0xb3b3b336b3b3b3edb3b3b3ff0000000000000000000000000000000000000000 ,
                        0x3255d62a000000000000000000000000b17d4a27b17d4af6b17d4a03b17d4ab7 ,
                        0xb17d4a6c0000000000000000000000000000000000000000000000003255d618 ,
                        0x3255d6333255d61bb3b3b3ffb3b3b3ffb3b3b33fb17d4a0cb17d4ab7b17d4aff ,
                        0xb17d4affb17d4a7bb3b3b327000000000000000000000000000000003255d67e ,
                        0x3255d6ff3255d68d00000000000000000000000000000000b17d4a87b17d4aff ,
                        0xb17d4affb17d4affb17d4a87000000000000000000000000000000003255d687 ,
                        0x3255d6ff3255d696b3b3b3ffb3b3b3ffb3b3b3ffb3b3b3eab3b3b32db17d4a8d ,
                        0xb17d4affb17d4affb17d4affb17d4a930000000000000000000000003255d693 ,
                        0x3255d6ff3255d6a5000000000000000000000000000000000000000000000000 ,
                        0xb17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a03000000003255d699 ,
                        0x3255d6ff3255d6b1000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa5000000003255d6a5 ,
                        0x3255d6ff3255d6ba000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a12b17d4acfb17d4affb17d4af0000000003255d6b1 ,
                        0x3255d6ff3255d6c3000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a000000003255d6b7 ,
                        0x3255d6ff3255d6d2000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000003255d6c3 ,
                        0x3255d6ff3255d6db000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =7058
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =7336
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =6
                    HoverForeShade =50.0
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =7
                    QuickStyleMask =-705
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12660
                    Top =7974
                    Width =1860
                    Height =278
                    FontSize =9
                    TabIndex =12
                    Name ="ScheduleCheck"
                    Caption ="Schedule Check"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =7974
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =8252
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =6
                    HoverForeShade =50.0
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =7
                    QuickStyleMask =-705
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12660
                    Top =7516
                    Width =1860
                    Height =278
                    FontSize =9
                    TabIndex =11
                    Name ="OpsReportExcel"
                    Caption ="OPS Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =7516
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =7794
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =6
                    HoverForeShade =50.0
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =7
                    QuickStyleMask =-705
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6780
                    Top =5280
                    Width =7380
                    Height =413
                    FontSize =12
                    FontWeight =700
                    TabIndex =14
                    Name ="SetLeanRelease"
                    Caption ="Set Lien Release for the Above Unit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000090a46869 ,
                        0x90a4686000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000090a4684e90a468fc ,
                        0x90a468f690a4682d000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000090a4683690a468f690a468ff ,
                        0x90a468ff90a468db90a4680f0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000090a4682d90a468ea90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468b40000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a4681b90a468e490a468ff90a468ff90a468ae ,
                        0x90a468ba90a468ff90a468ff90a4687e00000000000000000000000000000000 ,
                        0x000000000000000090a4680f90a468d290a468ff90a468ff90a468bd90a46806 ,
                        0x90a4681590a468e790a468ff90a468fc90a46848000000000000000000000000 ,
                        0x000000000000000090a4684290a468ff90a468ff90a468c990a4680f00000000 ,
                        0x0000000090a4683c90a468fc90a468ff90a468ed90a468210000000000000000 ,
                        0x00000000000000000000000090a4685790a468cc90a468150000000000000000 ,
                        0x000000000000000090a4687b90a468ff90a468ff90a468cc90a4680600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a468b790a468ff90a468ff90a4689c00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a4681590a468e190a468ff90a468ff90a46863 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000090a4683c90a468f990a468ff90a468f9 ,
                        0x90a4683300000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000090a4687590a468ff90a468ff ,
                        0x90a4689300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000090a468ab90a46887 ,
                        0x90a4680300000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =5280
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =5693
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =0
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =7020
                    Top =5040
                    Width =7200
                    Name ="Line84"
                    LayoutCachedLeft =7020
                    LayoutCachedTop =5040
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =5040
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =5940
                    Top =6120
                    Width =6360
                    Height =5325
                    FontSize =9
                    TabIndex =15
                    BoundColumn =5
                    Name ="LienReleaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientShortName AS Client,  [Schedule] & ' | ' & [UnitGroup] AS SchGrp, U"
                        "nitNum, VIN, LienRelease, AssetID  From vw_SixKeys  INNER JOIN Units on vw_SixKe"
                        "ys.AssetId=Units.UnitID  WHERE Not LienRelease is Null ORDER BY LienRelease Desc"
                        ";"
                    ColumnWidths ="1152;864;864;1872;720;0"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =6120
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =11445
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5940
                    Top =5880
                    Width =2685
                    Height =315
                    FontSize =10
                    FontWeight =700
                    Name ="Label86"
                    Caption ="Recent Lien Release entiries"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5940
                    LayoutCachedTop =5880
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =6195
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12660
                    Top =6600
                    Width =1860
                    Height =278
                    FontSize =9
                    TabIndex =16
                    Name ="LienReleasesRpt"
                    Caption =" Lien Releases"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12660
                    LayoutCachedTop =6600
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =6878
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BorderWidth =2
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ThemeFontIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =6
                    HoverForeShade =50.0
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    QuickStyle =7
                    QuickStyleMask =-705
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "TitleCheckin.cls"
