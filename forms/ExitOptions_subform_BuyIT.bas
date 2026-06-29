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
    Width =20922
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
            Height =1272
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
                    Left =36
                    Top =600
                    Width =1020
                    Height =300
                    FontSize =10
                    Name ="Label20"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =600
                    LayoutCachedWidth =1056
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
                    Left =7500
                    Top =60
                    Width =2556
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

                    LayoutCachedLeft =7500
                    LayoutCachedTop =60
                    LayoutCachedWidth =10056
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
                    Left =36
                    Top =60
                    Width =1020
                    Height =255
                    FontSize =8
                    FontWeight =700
                    Name ="send2Excel"
                    Caption ="Excel"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =36
                    LayoutCachedTop =60
                    LayoutCachedWidth =1056
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
                    Left =7500
                    Top =600
                    Width =2556
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
                    LayoutCachedLeft =7500
                    LayoutCachedTop =600
                    LayoutCachedWidth =10056
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
                    Left =10104
                    Top =600
                    Width =1392
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label454"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10104
                    LayoutCachedTop =600
                    LayoutCachedWidth =11496
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
                    Left =7404
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
                    LayoutCachedLeft =7404
                    LayoutCachedTop =600
                    LayoutCachedWidth =7404
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
                    Left =11568
                    Top =60
                    Width =1068
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

                    LayoutCachedLeft =11568
                    LayoutCachedTop =60
                    LayoutCachedWidth =12636
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
                    Left =19848
                    Top =360
                    Width =1068
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

                    LayoutCachedLeft =19848
                    LayoutCachedTop =360
                    LayoutCachedWidth =20916
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
                    Left =12708
                    Top =936
                    Width =3564
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =15
                    Name ="NotesCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =12708
                    LayoutCachedTop =936
                    LayoutCachedWidth =16272
                    LayoutCachedHeight =1236
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
                    Left =12708
                    Top =600
                    Width =3564
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label504"
                    Caption =" Note"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12708
                    LayoutCachedTop =600
                    LayoutCachedWidth =16272
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
                    Left =12708
                    Top =60
                    Width =1152
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

                    LayoutCachedLeft =12708
                    LayoutCachedTop =60
                    LayoutCachedWidth =13860
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
                    Left =19848
                    Top =936
                    Width =1068
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

                    LayoutCachedLeft =19848
                    LayoutCachedTop =936
                    LayoutCachedWidth =20916
                    LayoutCachedHeight =1236
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
                    Left =7404
                    Top =60
                    Width =0
                    Height =255
                    Name ="EmptyCell511"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =7404
                    LayoutCachedTop =60
                    LayoutCachedWidth =7404
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
                    Left =7404
                    Top =936
                    Width =0
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="Label1"
                    GroupTable =1

                    LayoutCachedLeft =7404
                    LayoutCachedTop =936
                    LayoutCachedWidth =7404
                    LayoutCachedHeight =1236
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
                    Left =7500
                    Top =936
                    Width =2556
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

                    LayoutCachedLeft =7500
                    LayoutCachedTop =936
                    LayoutCachedWidth =10056
                    LayoutCachedHeight =1236
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
                    Left =11568
                    Top =600
                    Width =1068
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Est. Delivery"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11568
                    LayoutCachedTop =600
                    LayoutCachedWidth =12636
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
                    Left =19848
                    Top =600
                    Width =1068
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
                    LayoutCachedLeft =19848
                    LayoutCachedTop =600
                    LayoutCachedWidth =20916
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
                    Left =7464
                    Top =60
                    Width =0
                    Height =255
                    Name ="EmptyCell540"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =7464
                    LayoutCachedTop =60
                    LayoutCachedWidth =7464
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1128
                    Top =600
                    Width =912
                    Height =300
                    FontSize =10
                    Name ="Label577"
                    Caption ="Unit#"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1128
                    LayoutCachedTop =600
                    LayoutCachedWidth =2040
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
                    Left =7464
                    Top =936
                    Width =0
                    Height =300
                    Name ="EmptyCell580"
                    GroupTable =1
                    LayoutCachedLeft =7464
                    LayoutCachedTop =936
                    LayoutCachedWidth =7464
                    LayoutCachedHeight =1236
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7464
                    Top =600
                    Width =0
                    Height =300
                    Name ="EmptyCell581"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7464
                    LayoutCachedTop =600
                    LayoutCachedWidth =7464
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10104
                    Top =60
                    Width =1392
                    Height =255
                    Name ="EmptyCell582"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =10104
                    LayoutCachedTop =60
                    LayoutCachedWidth =11496
                    LayoutCachedHeight =315
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =16308
                    Top =936
                    Width =1116
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

                    LayoutCachedLeft =16308
                    LayoutCachedTop =936
                    LayoutCachedWidth =17424
                    LayoutCachedHeight =1236
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
                    Left =16308
                    Top =60
                    Width =1116
                    Height =255
                    Name ="EmptyCell686"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =16308
                    LayoutCachedTop =60
                    LayoutCachedWidth =17424
                    LayoutCachedHeight =315
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16308
                    Top =600
                    Width =1116
                    Height =300
                    Name ="EmptyCell687"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16308
                    LayoutCachedTop =600
                    LayoutCachedWidth =17424
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
                    Left =1128
                    Top =60
                    Width =912
                    Height =255
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =1
                    Name ="txtCurrent"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =1128
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16308
                    Top =360
                    Width =1116
                    Name ="EmptyCell723"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16308
                    LayoutCachedTop =360
                    LayoutCachedWidth =17424
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
                    Left =36
                    Top =360
                    Width =7428
                    Height =240
                    FontSize =9
                    Name ="Label730"
                    Caption ="┌───────────────────── Non FA / Portfolio Equipment  ────────────────────┐"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =360
                    LayoutCachedWidth =7464
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
                    Left =7500
                    Top =360
                    Width =5136
                    Height =240
                    FontSize =9
                    ForeColor =8210719
                    Name ="Label732"
                    Caption ="┌──────────────── New Lease Equipment ───────────────┐"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7500
                    LayoutCachedTop =360
                    LayoutCachedWidth =12636
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
                    Left =3432
                    Top =936
                    Width =1248
                    Height =300
                    FontSize =9
                    TabIndex =10
                    Name ="AdjBookValueCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =3432
                    LayoutCachedTop =936
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1236
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
                    Left =3432
                    Top =600
                    Width =1248
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label741"
                    Caption ="Book Value  "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3432
                    LayoutCachedTop =600
                    LayoutCachedWidth =4680
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
                    Left =12708
                    Top =360
                    Width =3564
                    Name ="EmptyCell783"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12708
                    LayoutCachedTop =360
                    LayoutCachedWidth =16272
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10104
                    Top =936
                    Width =1392
                    Height =300
                    Name ="EmptyCell804"
                    GroupTable =1
                    LayoutCachedLeft =10104
                    LayoutCachedTop =936
                    LayoutCachedWidth =11496
                    LayoutCachedHeight =1236
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
                    Left =3432
                    Top =60
                    Width =1248
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

                    LayoutCachedLeft =3432
                    LayoutCachedTop =60
                    LayoutCachedWidth =4680
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
                    Left =1128
                    Top =936
                    Width =912
                    Height =300
                    Name ="EmptyCell816"
                    GroupTable =1
                    LayoutCachedLeft =1128
                    LayoutCachedTop =936
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1236
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =936
                    Width =1020
                    Height =300
                    Name ="EmptyCell817"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =936
                    LayoutCachedWidth =1056
                    LayoutCachedHeight =1236
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
                    Left =18768
                    Top =600
                    Width =1068
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
                    LayoutCachedLeft =18768
                    LayoutCachedTop =600
                    LayoutCachedWidth =19836
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
                    Left =18768
                    Top =936
                    Width =1068
                    Height =300
                    FontSize =10
                    ForeColor =8210719
                    Name ="Label826"
                    Caption ="Start Date"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18768
                    LayoutCachedTop =936
                    LayoutCachedWidth =19836
                    LayoutCachedHeight =1236
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
                    Left =13920
                    Top =60
                    Width =2352
                    Height =255
                    Name ="EmptyCell829"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =13920
                    LayoutCachedTop =60
                    LayoutCachedWidth =16272
                    LayoutCachedHeight =315
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4752
                    Top =600
                    Width =1272
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label831"
                    Caption ="Buy Out "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4752
                    LayoutCachedTop =600
                    LayoutCachedWidth =6024
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
                    Left =6096
                    Top =936
                    Width =1248
                    Height =300
                    FontSize =9
                    TabIndex =11
                    Name ="TargetSaleCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =6096
                    LayoutCachedTop =936
                    LayoutCachedWidth =7344
                    LayoutCachedHeight =1236
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
                    Left =6096
                    Top =600
                    Width =1248
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label842"
                    Caption ="Target "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6096
                    LayoutCachedTop =600
                    LayoutCachedWidth =7344
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
                    Left =6096
                    Top =60
                    Width =1248
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

                    LayoutCachedLeft =6096
                    LayoutCachedTop =60
                    LayoutCachedWidth =7344
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
                    Left =2112
                    Top =936
                    Width =1248
                    Height =300
                    FontSize =9
                    TabIndex =9
                    Name ="TermPenaltyCBO"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =2112
                    LayoutCachedTop =936
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1236
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
                    Left =2112
                    Top =600
                    Width =1248
                    Height =300
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label860"
                    Caption ="Term Penalty"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2112
                    LayoutCachedTop =600
                    LayoutCachedWidth =3360
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
                    Left =2112
                    Top =60
                    Width =1248
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

                    LayoutCachedLeft =2112
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
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
                    Left =4752
                    Top =60
                    Width =1272
                    Height =255
                    Name ="EmptyCell872"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =4752
                    LayoutCachedTop =60
                    LayoutCachedWidth =6024
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
                    Left =4752
                    Top =936
                    Width =1272
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =4754549
                    ForeColor =26112
                    Name ="Label874"
                    Caption ="Cost"
                    GroupTable =1
                    LayoutCachedLeft =4752
                    LayoutCachedTop =936
                    LayoutCachedWidth =6024
                    LayoutCachedHeight =1236
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
                    Left =11568
                    Top =936
                    Width =1068
                    Height =300
                    FontSize =9
                    TabIndex =14
                    Name ="StartOrPurposedDateCBO"
                    Format ="Short Date"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =11568
                    LayoutCachedTop =936
                    LayoutCachedWidth =12636
                    LayoutCachedHeight =1236
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
                    Left =18768
                    Top =60
                    Width =1068
                    Height =255
                    Name ="EmptyCell918"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18768
                    LayoutCachedTop =60
                    LayoutCachedWidth =19836
                    LayoutCachedHeight =315
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19848
                    Top =60
                    Width =1068
                    Height =255
                    Name ="EmptyCell921"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =19848
                    LayoutCachedTop =60
                    LayoutCachedWidth =20916
                    LayoutCachedHeight =315
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18768
                    Top =360
                    Width =1068
                    Name ="EmptyCell922"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18768
                    LayoutCachedTop =360
                    LayoutCachedWidth =19836
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18456
                    Top =60
                    Width =300
                    Height =255
                    Name ="EmptyCell925"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18456
                    LayoutCachedTop =60
                    LayoutCachedWidth =18756
                    LayoutCachedHeight =315
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18456
                    Top =360
                    Width =300
                    Name ="EmptyCell926"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18456
                    LayoutCachedTop =360
                    LayoutCachedWidth =18756
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18456
                    Top =600
                    Width =300
                    Height =300
                    Name ="EmptyCell927"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18456
                    LayoutCachedTop =600
                    LayoutCachedWidth =18756
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18456
                    Top =936
                    Width =300
                    Height =300
                    Name ="EmptyCell928"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18456
                    LayoutCachedTop =936
                    LayoutCachedWidth =18756
                    LayoutCachedHeight =1236
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17424
                    Top =60
                    Width =1032
                    Height =255
                    Name ="EmptyCell930"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =17424
                    LayoutCachedTop =60
                    LayoutCachedWidth =18456
                    LayoutCachedHeight =315
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17424
                    Top =360
                    Width =1032
                    Name ="EmptyCell931"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17424
                    LayoutCachedTop =360
                    LayoutCachedWidth =18456
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17424
                    Top =600
                    Width =1032
                    Height =300
                    Name ="EmptyCell932"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17424
                    LayoutCachedTop =600
                    LayoutCachedWidth =18456
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17424
                    Top =936
                    Width =1032
                    Height =300
                    Name ="EmptyCell933"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =17424
                    LayoutCachedTop =936
                    LayoutCachedWidth =18456
                    LayoutCachedHeight =1236
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
                    Left =18456
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

                    LayoutCachedLeft =18456
                    LayoutCachedWidth =18756
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
                    Left =7404
                    Width =0
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="OptionTypeID"
                    ControlSource ="OptionTypeID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7404
                    LayoutCachedWidth =7404
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
                    Left =19848
                    Width =1068
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

                    LayoutCachedLeft =19848
                    LayoutCachedWidth =20916
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
                    Left =36
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

                    LayoutCachedLeft =36
                    LayoutCachedWidth =1056
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
                    Left =11568
                    Width =1068
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

                    LayoutCachedLeft =11568
                    LayoutCachedWidth =12636
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
                    Left =7464
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

                    LayoutCachedLeft =7464
                    LayoutCachedWidth =7464
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
                    Left =12708
                    Width =3564
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

                    LayoutCachedLeft =12708
                    LayoutCachedWidth =16272
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
                    Left =10104
                    Width =1392
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

                    LayoutCachedLeft =10104
                    LayoutCachedWidth =11496
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
                    Left =7500
                    Width =2556
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

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =10056
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
                    Left =16308
                    Width =1116
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

                    LayoutCachedLeft =16308
                    LayoutCachedWidth =17424
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
                    Left =17424
                    Width =1032
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

                    LayoutCachedLeft =17424
                    LayoutCachedWidth =18456
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
                    Left =1128
                    Width =912
                    Height =270
                    FontSize =10
                    TabIndex =1
                    BackColor =16317438
                    Name ="Text575"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1128
                    LayoutCachedWidth =2040
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
                    Left =3432
                    Width =1248
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

                    LayoutCachedLeft =3432
                    LayoutCachedWidth =4680
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
                    Left =18768
                    Width =1068
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

                    LayoutCachedLeft =18768
                    LayoutCachedWidth =19836
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
                    Left =4752
                    Width =1272
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

                    LayoutCachedLeft =4752
                    LayoutCachedWidth =6024
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
                    Left =6096
                    Width =1248
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

                    LayoutCachedLeft =6096
                    LayoutCachedWidth =7344
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
                    Left =2112
                    Width =1248
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

                    LayoutCachedLeft =2112
                    LayoutCachedWidth =3360
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
