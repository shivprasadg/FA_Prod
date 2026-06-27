Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20880
    DatasheetFontHeight =11
    ItemSuffix =935
    Left =7590
    Top =6780
    Right =26475
    Bottom =18900
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x3c47465b8cd0e540
    End
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1260
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =599
                    Width =20880
                    Height =300
                    Name ="Box47"
                    LayoutCachedTop =599
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =899
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Top =599
                    Width =7320
                    Height =300
                    BorderColor =1643706
                    Name ="Box815"
                    LayoutCachedTop =599
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =899
                    BackThemeColorIndex =6
                    BackShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =95
                    Top =900
                    Width =20880
                    Height =360
                    BackColor =15527148
                    Name ="Box729"
                    LayoutCachedTop =900
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =30
                    Top =600
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label20"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =600
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7425
                    Top =60
                    Width =2550
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="ReplcedBySubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the GROUP that holds the units that will replace this unit. for all listed u"
                        "nits"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =15
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7425
                    LayoutCachedTop =60
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderColor =8210719
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =30
                    Top =60
                    Width =1020
                    Height =255
                    FontSize =8
                    FontWeight =700
                    Name ="send2Excel"
                    Caption ="Excel"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =30
                    LayoutCachedTop =60
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =315
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7425
                    Top =600
                    Width =2550
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Schedule | Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =7425
                    LayoutCachedTop =600
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10020
                    Top =600
                    Width =1395
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label454"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10020
                    LayoutCachedTop =600
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7335
                    Top =600
                    Width =0
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label478"
                    Caption ="OptionID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7335
                    LayoutCachedTop =600
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11475
                    Top =60
                    Width =1065
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="StartOrPurposedDateCBOSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11475
                    LayoutCachedTop =60
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =315
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderColor =8210719
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =95
                    Left =19710
                    Top =360
                    Width =1065
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="ActualDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the Actual ExchangeIT date for all listed units"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19710
                    LayoutCachedTop =360
                    LayoutCachedWidth =20775
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderColor =8210719
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12600
                    Top =930
                    Width =3570
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =15
                    Name ="NotesCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =12600
                    LayoutCachedTop =930
                    LayoutCachedWidth =16170
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =12600
                    Top =600
                    Width =3570
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label504"
                    Caption =" Note"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12600
                    LayoutCachedTop =600
                    LayoutCachedWidth =16170
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12600
                    Top =60
                    Width =1155
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="NotesSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12600
                    LayoutCachedTop =60
                    LayoutCachedWidth =13755
                    LayoutCachedHeight =315
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderColor =8210719
                    HoverColor =8210719
                    PressedThemeColorIndex =2
                    PressedTint =60.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19710
                    Top =930
                    Width =1065
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =17
                    Name ="EndOrActualDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="Set the Actual ExchangeIT date for all listed units"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =19710
                    LayoutCachedTop =930
                    LayoutCachedWidth =20775
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7335
                    Top =60
                    Width =0
                    Height =255
                    Name ="EmptyCell511"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =7335
                    LayoutCachedTop =60
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7335
                    Top =930
                    Width =0
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="Label1"
                    GroupTable =1

                    LayoutCachedLeft =7335
                    LayoutCachedTop =930
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3456
                    Left =7425
                    Top =930
                    Width =2550
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =13
                    Name ="ReplacedByGroupCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.SGrpID, [Schedule] & \" | \" & [UnitGroup] AS [Schedule-Group]"
                        ", Count(vw_SixKeys.UnitID) AS Units, vw_SixKeys.SchID FROM vw_SixKeys LEFT JOIN "
                        "UnitExitOption ON vw_SixKeys.UnitID = UnitExitOption.ReplacedByUnitID WHERE (((v"
                        "w_SixKeys.Status)<>\"T\") AND ((vw_SixKeys.clientID)=[Forms]![EXITOptions]![Clie"
                        "ntList]) AND ((UnitExitOption.ReplacedByUnitID) Is Null) AND ((vw_SixKeys.MLOrig"
                        ")=\"FA\")) GROUP BY vw_SixKeys.SGrpID, [Schedule] & \" | \" & [UnitGroup], vw_Si"
                        "xKeys.SchID ORDER BY vw_SixKeys.SchID;"
                    ColumnWidths ="0;2448;432"
                    Tag ="Medit"
                    ControlTipText ="Set the GROUP that holds the units that will replace this unit. for all listed u"
                        "nits"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =15

                    LayoutCachedLeft =7425
                    LayoutCachedTop =930
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =11475
                    Top =600
                    Width =1065
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Est. Delivery"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11475
                    LayoutCachedTop =600
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =19710
                    Top =600
                    Width =1065
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Text520"
                    Caption ="In-Serv Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19710
                    LayoutCachedTop =600
                    LayoutCachedWidth =20775
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7395
                    Top =60
                    Width =0
                    Height =255
                    Name ="EmptyCell540"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =7395
                    LayoutCachedTop =60
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1110
                    Top =600
                    Width =915
                    Height =300
                    FontSize =10
                    Name ="Label577"
                    Caption ="Unit#"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1110
                    LayoutCachedTop =600
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7395
                    Top =930
                    Width =0
                    Height =300
                    Name ="EmptyCell580"
                    GroupTable =1
                    LayoutCachedLeft =7395
                    LayoutCachedTop =930
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7395
                    Top =600
                    Width =0
                    Height =300
                    Name ="EmptyCell581"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7395
                    LayoutCachedTop =600
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10020
                    Top =60
                    Width =1395
                    Height =255
                    Name ="EmptyCell582"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =10020
                    LayoutCachedTop =60
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =315
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =16200
                    Top =930
                    Width =1110
                    Height =300
                    FontSize =8
                    TabIndex =16
                    Name ="SaveUpdateMe"
                    Caption ="Save/Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save Update Changes"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =16200
                    LayoutCachedTop =930
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderColor =16777215
                    HoverColor =12428174
                    PressedColor =8210719
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =16200
                    Top =60
                    Width =1110
                    Height =255
                    Name ="EmptyCell686"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =16200
                    LayoutCachedTop =60
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =315
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16200
                    Top =600
                    Width =1110
                    Height =300
                    Name ="EmptyCell687"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16200
                    LayoutCachedTop =600
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1110
                    Top =60
                    Width =915
                    Height =255
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =1
                    Name ="txtCurrent"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =1110
                    LayoutCachedTop =60
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16200
                    Top =360
                    Width =1110
                    Name ="EmptyCell723"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16200
                    LayoutCachedTop =360
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =30
                    Top =360
                    Width =7365
                    Height =240
                    FontSize =9
                    Name ="Label730"
                    Caption ="┌───────────────────── Non FA / Portfolio Equipment  ────────────────────┐"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =360
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7425
                    Top =360
                    Width =5115
                    Height =240
                    FontSize =9
                    ForeColor =8210719
                    Name ="Label732"
                    Caption ="┌──────────────── New Lease Equipment ───────────────┐"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7425
                    LayoutCachedTop =360
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3390
                    Top =930
                    Width =1245
                    Height =300
                    FontSize =9
                    TabIndex =10
                    Name ="AdjBookValueCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =3390
                    LayoutCachedTop =930
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3390
                    Top =600
                    Width =1245
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label741"
                    Caption ="Book Value  "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3390
                    LayoutCachedTop =600
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12600
                    Top =360
                    Width =3570
                    Name ="EmptyCell783"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12600
                    LayoutCachedTop =360
                    LayoutCachedWidth =16170
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10020
                    Top =930
                    Width =1395
                    Height =300
                    Name ="EmptyCell804"
                    GroupTable =1
                    LayoutCachedLeft =10020
                    LayoutCachedTop =930
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3390
                    Top =60
                    Width =1245
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="AdjBookValueSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the Last Rent Date for all listed units"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3390
                    LayoutCachedTop =60
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =4754549
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin EmptyCell
                    Left =1110
                    Top =930
                    Width =915
                    Height =300
                    Name ="EmptyCell816"
                    GroupTable =1
                    LayoutCachedLeft =1110
                    LayoutCachedTop =930
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =30
                    Top =930
                    Width =1020
                    Height =300
                    Name ="EmptyCell817"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =930
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =18645
                    Top =600
                    Width =1065
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label824"
                    Caption ="Per Diem"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18645
                    LayoutCachedTop =600
                    LayoutCachedWidth =19710
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =18645
                    Top =930
                    Width =1065
                    Height =300
                    FontSize =10
                    ForeColor =8210719
                    Name ="Label826"
                    Caption ="Start Date"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18645
                    LayoutCachedTop =930
                    LayoutCachedWidth =19710
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =13815
                    Top =60
                    Width =2355
                    Height =255
                    Name ="EmptyCell829"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =13815
                    LayoutCachedTop =60
                    LayoutCachedWidth =16170
                    LayoutCachedHeight =315
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4695
                    Top =600
                    Width =1275
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label831"
                    Caption ="Buy Out "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4695
                    LayoutCachedTop =600
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6030
                    Top =930
                    Width =1245
                    Height =300
                    FontSize =9
                    TabIndex =11
                    Name ="TargetSaleCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =6030
                    LayoutCachedTop =930
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6030
                    Top =600
                    Width =1245
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label842"
                    Caption ="Target "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6030
                    LayoutCachedTop =600
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6030
                    Top =60
                    Width =1245
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="TargetSaleSubmit"
                    Caption ="▼Submit"
                    ControlTipText ="Set the Last Rent Date for all listed units"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6030
                    LayoutCachedTop =60
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =315
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =4754549
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2085
                    Top =930
                    Width =1245
                    Height =300
                    FontSize =9
                    TabIndex =9
                    Name ="TermPenaltyCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =2085
                    LayoutCachedTop =930
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2085
                    Top =600
                    Width =1245
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label860"
                    Caption ="Term Penalty"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2085
                    LayoutCachedTop =600
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2085
                    Top =60
                    Width =1245
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="TermPenaltySubmit"
                    Caption ="▼Submit"
                    ControlTipText ="Set the Last Rent Date for all listed units"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2085
                    LayoutCachedTop =60
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =4754549
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin EmptyCell
                    Left =4695
                    Top =60
                    Width =1275
                    Height =255
                    Name ="EmptyCell872"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =4695
                    LayoutCachedTop =60
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4695
                    Top =930
                    Width =1275
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =4754549
                    ForeColor =26112
                    Name ="Label874"
                    Caption ="Cost"
                    GroupTable =1
                    LayoutCachedLeft =4695
                    LayoutCachedTop =930
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11475
                    Top =930
                    Width =1065
                    Height =300
                    FontSize =9
                    TabIndex =14
                    Name ="StartOrPurposedDateCBO"
                    Format ="Short Date"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =11475
                    LayoutCachedTop =930
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18645
                    Top =60
                    Width =1065
                    Height =255
                    Name ="EmptyCell918"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18645
                    LayoutCachedTop =60
                    LayoutCachedWidth =19710
                    LayoutCachedHeight =315
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19710
                    Top =60
                    Width =1065
                    Height =255
                    Name ="EmptyCell921"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =19710
                    LayoutCachedTop =60
                    LayoutCachedWidth =20775
                    LayoutCachedHeight =315
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18645
                    Top =360
                    Width =1065
                    Name ="EmptyCell922"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18645
                    LayoutCachedTop =360
                    LayoutCachedWidth =19710
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18345
                    Top =60
                    Width =300
                    Height =255
                    Name ="EmptyCell925"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18345
                    LayoutCachedTop =60
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =315
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18345
                    Top =360
                    Width =300
                    Name ="EmptyCell926"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18345
                    LayoutCachedTop =360
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18345
                    Top =600
                    Width =300
                    Height =300
                    Name ="EmptyCell927"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18345
                    LayoutCachedTop =600
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18345
                    Top =930
                    Width =300
                    Height =300
                    Name ="EmptyCell928"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18345
                    LayoutCachedTop =930
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17310
                    Top =60
                    Width =1035
                    Height =255
                    Name ="EmptyCell930"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =17310
                    LayoutCachedTop =60
                    LayoutCachedWidth =18345
                    LayoutCachedHeight =315
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17310
                    Top =360
                    Width =1035
                    Name ="EmptyCell931"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17310
                    LayoutCachedTop =360
                    LayoutCachedWidth =18345
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17310
                    Top =600
                    Width =1035
                    Height =300
                    Name ="EmptyCell932"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17310
                    LayoutCachedTop =600
                    LayoutCachedWidth =18345
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17310
                    Top =930
                    Width =1035
                    Height =300
                    Name ="EmptyCell933"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =17310
                    LayoutCachedTop =930
                    LayoutCachedWidth =18345
                    LayoutCachedHeight =1230
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =359
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =18345
                    Width =300
                    Height =270
                    FontSize =10
                    TabIndex =14
                    ForeColor =1643706
                    Name ="DeleteMe"
                    Caption ="Õ"
                    OnClick ="[Event Procedure]"
                    FontName ="Wingdings"
                    ControlTipText ="Delete This Unit's Option"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18345
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =1643706
                    PressedThemeColorIndex =9
                    PressedShade =75.0
                    PressedForeColor =1643706
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7335
                    Width =0
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="OptionTypeID"
                    ControlSource ="OptionTypeID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7335
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19710
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =16
                    Name ="EndOrActualDate"
                    ControlSource ="EndOrActualDate"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Accepted Date of the NEW Unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19710
                    LayoutCachedWidth =20775
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Width =1020
                    Height =270
                    FontSize =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =300
                    Width =20880
                    Height =59
                    TabIndex =17
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =300
                    LayoutCachedWidth =20880
                    LayoutCachedHeight =359
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11475
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =10
                    Name ="StartOrPurposedDate"
                    ControlSource ="StartOrPurposedDate"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Estimated Delivery Date of the NEW Unit"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11475
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7395
                    Width =0
                    Height =270
                    FontSize =10
                    TabIndex =7
                    BackColor =16317438
                    Name ="ReplacedByScheduleID"
                    ControlSource ="ReplacedByScheduleID"
                    Format ="Standard"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7395
                    LayoutCachedWidth =7395
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Width =3570
                    Height =270
                    FontSize =9
                    TabIndex =11
                    Name ="Notes"
                    ControlSource ="Notes"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="notes..."
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12600
                    LayoutCachedWidth =16170
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =8640
                    Left =10020
                    Width =1395
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="ReplacedByUnitID"
                    ControlSource ="ReplacedByUnitID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.UnitID AS AssetID, vw_SixKeys.Unitnum AS [Unit#], [Myear] & \""
                        " \" & [Make] & \" \" & [Model] & \" \" & [Type] AS [Desc], vw_SixKeys.SchID, vw_"
                        "SixKeys.Vin AS VIN FROM vw_SixKeys LEFT JOIN UnitExitOption ON vw_SixKeys.UnitID"
                        " = UnitExitOption.ReplacedByUnitID WHERE (((vw_SixKeys.SGrpID)=[Forms]![EXITOpti"
                        "ons]![ExitOptions_subform_BuyIT].[Form]![ReplacedByGroupID]) AND ((UnitExitOptio"
                        "n.ReplacedByUnitID) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\"));"
                    ColumnWidths ="720;720;3888;0;2592"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="This is the exact unit the client chose to replace this unit with, Its Delivery "
                        "and ISD dates will be unsed"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10020
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3456
                    Left =7425
                    Width =2550
                    Height =270
                    FontSize =10
                    TabIndex =8
                    BackColor =16317438
                    Name ="ReplacedByGroupID"
                    ControlSource ="ReplacedByGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.SGrpID, [Schedule] & \" | \" & [UnitGroup] AS [Schedule-Group]"
                        ", Count(vw_SixKeys.UnitID) AS Units, vw_SixKeys.SchID FROM vw_SixKeys LEFT JOIN "
                        "UnitExitOption ON vw_SixKeys.UnitID = UnitExitOption.ReplacedByUnitID WHERE (((v"
                        "w_SixKeys.Status)<>\"T\") AND ((vw_SixKeys.clientID)=[Forms]![EXITOptions]![Clie"
                        "ntList]) AND ((vw_SixKeys.MLOrig)=\"FA\")) GROUP BY vw_SixKeys.SGrpID, [Schedule"
                        "] & \" | \" & [UnitGroup], vw_SixKeys.SchID ORDER BY vw_SixKeys.SchID;"
                    ColumnWidths ="0;2448;432"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="This is the GROUP that holds the units that will replace this unit."
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0

                    LayoutCachedLeft =7425
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16200
                    Width =1110
                    Height =270
                    FontSize =8
                    TabIndex =12
                    Name ="UserName"
                    ControlSource ="UserName"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16200
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17310
                    Width =1035
                    Height =270
                    FontSize =8
                    TabIndex =13
                    Name ="DatePosted"
                    ControlSource ="DatePosted"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17310
                    LayoutCachedWidth =18345
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1110
                    Width =915
                    Height =270
                    FontSize =10
                    TabIndex =1
                    BackColor =16317438
                    Name ="Text575"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3390
                    Width =1245
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="AdjBookValue"
                    ControlSource ="AdjBookValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3390
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18645
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =15
                    Name ="UnitPDStartDate"
                    ControlSource ="UnitPDStartDate"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Per Diem Start Date of the NEW Unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18645
                    LayoutCachedWidth =19710
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4695
                    Width =1275
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="BuyOutCost"
                    ControlSource ="BuyOutCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4695
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6030
                    Width =1245
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="TargetSale"
                    ControlSource ="TargetSale"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6030
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2085
                    Width =1245
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="TermPenalty"
                    ControlSource ="TermPenalty"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2085
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14385
                    Width =29
                    Height =270
                    FontWeight =700
                    TabIndex =2
                    Name ="SchNo"

                    LayoutCachedLeft =14385
                    LayoutCachedWidth =14414
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =14265
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =14265
                    LayoutCachedWidth =14294
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14115
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =14115
                    LayoutCachedWidth =14144
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
            End
        End
    End
End
CodeBehindForm
' See "ExitOptions_subform_BuyIT.cls"
