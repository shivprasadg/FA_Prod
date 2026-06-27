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
    Width =20948
    DatasheetFontHeight =11
    ItemSuffix =858
    Left =6165
    Top =4905
    Right =26205
    Bottom =9420
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x98925c478fbce440
    End
    RecordSource ="ExitOptionsQryExchangeIT"
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
            Height =1320
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
                    Width =4080
                    Height =300
                    BorderColor =1643706
                    Name ="Box815"
                    LayoutCachedTop =599
                    LayoutCachedWidth =4080
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
                    Width =960
                    Height =285
                    FontSize =10
                    Name ="Label20"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =600
                    LayoutCachedWidth =990
                    LayoutCachedHeight =885
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
                    Left =4140
                    Top =60
                    Width =2520
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
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

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
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
                    Width =960
                    Height =255
                    FontSize =8
                    FontWeight =700
                    Name ="send2Excel"
                    Caption ="Excel"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =30
                    LayoutCachedTop =60
                    LayoutCachedWidth =990
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
                    Left =4140
                    Top =600
                    Width =2520
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Schedule | Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =4140
                    LayoutCachedTop =600
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =885
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
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6705
                    Top =600
                    Width =1395
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label454"
                    Caption ="AssetID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6705
                    LayoutCachedTop =600
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4035
                    Top =600
                    Width =0
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label478"
                    Caption ="OptionID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4035
                    LayoutCachedTop =600
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =885
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =9300
                    Top =60
                    Width =1080
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="ActualDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the Actual ExchangeIT date for all listed units"
                    GroupTable =1
                    RightPadding =0
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9300
                    LayoutCachedTop =60
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =315
                    ColumnStart =9
                    ColumnEnd =9
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
                    Left =10410
                    Top =915
                    Width =4890
                    Height =300
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =11
                    Name ="NotesCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =10410
                    LayoutCachedTop =915
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =11
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
                    Left =10410
                    Top =600
                    Width =4890
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label504"
                    Caption =" ExchangeIT Note"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10410
                    LayoutCachedTop =600
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10410
                    Top =60
                    Width =1155
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="NotesSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10410
                    LayoutCachedTop =60
                    LayoutCachedWidth =11565
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
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9300
                    Top =915
                    Width =1080
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =10
                    Name ="EndOrActualDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="Set the Actual ExchangeIT date for all listed units"
                    GroupTable =1
                    RightPadding =0

                    LayoutCachedLeft =9300
                    LayoutCachedTop =915
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4035
                    Top =60
                    Width =0
                    Height =255
                    Name ="EmptyCell511"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =4035
                    LayoutCachedTop =60
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4035
                    Top =915
                    Width =0
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =8
                    BorderColor =0
                    Name ="Label1"
                    GroupTable =1

                    LayoutCachedLeft =4035
                    LayoutCachedTop =915
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =7200
                    Left =4140
                    Top =915
                    Width =2520
                    Height =300
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =9
                    Name ="ReplacedByGroupCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.SGrpID, [ClientShortName] & \" | \" & [Schedule] & \" | \" & ["
                        "UnitGroup] AS [Schedule-Group], Count(vw_SixKeys.UnitID) AS Units, vw_SixKeys.Sc"
                        "hID FROM vw_SixKeys LEFT JOIN UnitExitOption ON vw_SixKeys.UnitID = UnitExitOpti"
                        "on.ReplacedByUnitID WHERE (((vw_SixKeys.Status)<>\"T\") AND ((UnitExitOption.Rep"
                        "lacedByUnitID) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\") AND ((vw_SixKeys.Client"
                        "GroupId)=[Forms]![EXITOptions]![ClientGroupID])) GROUP BY vw_SixKeys.SGrpID, [Cl"
                        "ientShortName] & \" | \" & [Schedule] & \" | \" & [UnitGroup], vw_SixKeys.SchID "
                        "ORDER BY [ClientShortName] & \" | \" & [Schedule] & \" | \" & [UnitGroup];"
                    ColumnWidths ="0;2448;432"
                    Tag ="Medit"
                    ControlTipText ="Set the GROUP that holds the units that will replace this unit. for all listed u"
                        "nits"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =15

                    LayoutCachedLeft =4140
                    LayoutCachedTop =915
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8160
                    Top =600
                    Width =1080
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Est. Delivery"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8160
                    LayoutCachedTop =600
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =885
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
                    Left =9300
                    Top =600
                    Width =1080
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Text520"
                    Caption ="Accepted"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9300
                    LayoutCachedTop =600
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =885
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
                Begin EmptyCell
                    Left =4095
                    Top =60
                    Width =15
                    Height =255
                    Name ="EmptyCell540"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =4095
                    LayoutCachedTop =60
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =315
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1050
                    Top =600
                    Width =915
                    Height =285
                    FontSize =10
                    Name ="Label577"
                    Caption ="Unit#"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1050
                    LayoutCachedTop =600
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =885
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
                    Left =4095
                    Top =915
                    Width =15
                    Height =300
                    Name ="EmptyCell580"
                    GroupTable =1
                    LayoutCachedLeft =4095
                    LayoutCachedTop =915
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4095
                    Top =600
                    Width =15
                    Height =285
                    Name ="EmptyCell581"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4095
                    LayoutCachedTop =600
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6705
                    Top =60
                    Width =1395
                    Height =255
                    Name ="EmptyCell582"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =6705
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =16335
                    Top =915
                    Width =1110
                    Height =300
                    FontSize =8
                    TabIndex =12
                    Name ="SaveUpdateMe"
                    Caption ="Save/Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save Update Changes"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =16335
                    LayoutCachedTop =915
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =15840
                    Top =60
                    Width =465
                    Height =255
                    Name ="EmptyCell634"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =15840
                    LayoutCachedTop =60
                    LayoutCachedWidth =16305
                    LayoutCachedHeight =315
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15840
                    Top =600
                    Width =465
                    Height =285
                    Name ="EmptyCell635"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15840
                    LayoutCachedTop =600
                    LayoutCachedWidth =16305
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15840
                    Top =915
                    Width =465
                    Height =300
                    Name ="EmptyCell636"
                    GroupTable =1
                    LayoutCachedLeft =15840
                    LayoutCachedTop =915
                    LayoutCachedWidth =16305
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15345
                    Top =60
                    Width =465
                    Height =255
                    Name ="EmptyCell638"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =15345
                    LayoutCachedTop =60
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =315
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15345
                    Top =600
                    Width =465
                    Height =285
                    Name ="EmptyCell639"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15345
                    LayoutCachedTop =600
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15345
                    Top =915
                    Width =465
                    Height =300
                    Name ="EmptyCell640"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =15345
                    LayoutCachedTop =915
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16335
                    Top =60
                    Width =1110
                    Height =255
                    Name ="EmptyCell686"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =16335
                    LayoutCachedTop =60
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =315
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16335
                    Top =600
                    Width =1110
                    Height =285
                    Name ="EmptyCell687"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16335
                    LayoutCachedTop =600
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17445
                    Top =60
                    Width =1035
                    Height =255
                    Name ="EmptyCell692"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =17445
                    LayoutCachedTop =60
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =315
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17445
                    Top =600
                    Width =1035
                    Height =285
                    Name ="EmptyCell693"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17445
                    LayoutCachedTop =600
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17445
                    Top =915
                    Width =1035
                    Height =300
                    Name ="EmptyCell694"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =17445
                    LayoutCachedTop =915
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18480
                    Top =60
                    Width =300
                    Height =255
                    Name ="EmptyCell701"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18480
                    LayoutCachedTop =60
                    LayoutCachedWidth =18780
                    LayoutCachedHeight =315
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18480
                    Top =600
                    Width =300
                    Height =285
                    Name ="EmptyCell702"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18480
                    LayoutCachedTop =600
                    LayoutCachedWidth =18780
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18480
                    Top =915
                    Width =300
                    Height =300
                    Name ="EmptyCell703"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18480
                    LayoutCachedTop =915
                    LayoutCachedWidth =18780
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18780
                    Top =60
                    Width =2115
                    Height =255
                    Name ="EmptyCell707"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =18780
                    LayoutCachedTop =60
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =315
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18780
                    Top =600
                    Width =2115
                    Height =285
                    Name ="EmptyCell708"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18780
                    LayoutCachedTop =600
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18780
                    Top =915
                    Width =2115
                    Height =300
                    Name ="EmptyCell709"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =18780
                    LayoutCachedTop =915
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1050
                    Top =60
                    Width =915
                    Height =255
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =1
                    Name ="txtCurrent"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =1050
                    LayoutCachedTop =60
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15345
                    Top =360
                    Width =960
                    Name ="EmptyCell721"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15345
                    LayoutCachedTop =360
                    LayoutCachedWidth =16305
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =16335
                    Top =360
                    Width =1110
                    Name ="EmptyCell723"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16335
                    LayoutCachedTop =360
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17445
                    Top =360
                    Width =1035
                    Name ="EmptyCell724"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17445
                    LayoutCachedTop =360
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18480
                    Top =360
                    Width =300
                    Name ="EmptyCell725"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18480
                    LayoutCachedTop =360
                    LayoutCachedWidth =18780
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =18780
                    Top =360
                    Width =2115
                    Name ="EmptyCell726"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18780
                    LayoutCachedTop =360
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =30
                    Top =360
                    Width =4080
                    Height =240
                    FontSize =9
                    Name ="Label730"
                    Caption ="┌────── Original Equipment  ──────┐"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =360
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =5
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
                    Left =4140
                    Top =360
                    Width =6240
                    Height =240
                    FontSize =9
                    ForeColor =8210719
                    Name ="Label732"
                    Caption ="┌───────────────────── EXchangeIT Equipment ────────────────────┐"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4140
                    LayoutCachedTop =360
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2025
                    Top =915
                    Width =945
                    Height =300
                    FontSize =9
                    TabIndex =7
                    Name ="FinalRentDateCBO"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1

                    LayoutCachedLeft =2025
                    LayoutCachedTop =915
                    LayoutCachedWidth =2970
                    LayoutCachedHeight =1215
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
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2025
                    Top =600
                    Width =945
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label741"
                    Caption ="Final Rent "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2025
                    LayoutCachedTop =600
                    LayoutCachedWidth =2970
                    LayoutCachedHeight =885
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
                Begin EmptyCell
                    Left =10410
                    Top =360
                    Width =4890
                    Name ="EmptyCell783"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10410
                    LayoutCachedTop =360
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6705
                    Top =915
                    Width =1395
                    Height =300
                    Name ="EmptyCell804"
                    GroupTable =1
                    LayoutCachedLeft =6705
                    LayoutCachedTop =915
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2025
                    Top =60
                    Width =945
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="FinalRentDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the Last Rent Date for all listed units"
                    GroupTable =1
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2025
                    LayoutCachedTop =60
                    LayoutCachedWidth =2970
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
                    Left =1050
                    Top =915
                    Width =915
                    Height =300
                    Name ="EmptyCell816"
                    GroupTable =1
                    LayoutCachedLeft =1050
                    LayoutCachedTop =915
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =30
                    Top =915
                    Width =960
                    Height =300
                    Name ="EmptyCell817"
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =915
                    LayoutCachedWidth =990
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =20895
                    Top =600
                    Width =45
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label824"
                    Caption ="Per Diem"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20895
                    LayoutCachedTop =600
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8160
                    Top =915
                    Width =1080
                    Height =300
                    Name ="EmptyCell828"
                    GroupTable =1
                    LayoutCachedLeft =8160
                    LayoutCachedTop =915
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11625
                    Top =60
                    Width =3675
                    Height =255
                    Name ="EmptyCell829"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =11625
                    LayoutCachedTop =60
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =315
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8160
                    Top =60
                    Width =1080
                    Height =255
                    Name ="EmptyCell830"
                    GroupTable =1
                    BottomPadding =45
                    LayoutCachedLeft =8160
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =20895
                    Top =60
                    Width =45
                    Height =255
                    Name ="EmptyCell834"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =20895
                    LayoutCachedTop =60
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =315
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =20895
                    Top =360
                    Width =45
                    Name ="EmptyCell835"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20895
                    LayoutCachedTop =360
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =20895
                    Top =915
                    Width =45
                    Height =300
                    Name ="EmptyCell837"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =20895
                    LayoutCachedTop =915
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3030
                    Top =915
                    Width =945
                    Height =300
                    Name ="EmptyCell843"
                    GroupTable =1
                    LayoutCachedLeft =3030
                    LayoutCachedTop =915
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3030
                    Top =600
                    Width =945
                    Height =285
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label845"
                    Caption ="Proj Eol"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3030
                    LayoutCachedTop =600
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =885
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
                    Left =15330
                    Top =60
                    Width =14
                    Height =255
                    Name ="EmptyCell851"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    BottomPadding =45
                    LayoutCachedLeft =15330
                    LayoutCachedTop =60
                    LayoutCachedWidth =15344
                    LayoutCachedHeight =315
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15330
                    Top =360
                    Width =14
                    Name ="EmptyCell852"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15330
                    LayoutCachedTop =360
                    LayoutCachedWidth =15344
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15330
                    Top =600
                    Width =14
                    Height =285
                    Name ="EmptyCell853"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15330
                    LayoutCachedTop =600
                    LayoutCachedWidth =15344
                    LayoutCachedHeight =885
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =15330
                    Top =915
                    Width =14
                    Height =300
                    Name ="EmptyCell854"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =15330
                    LayoutCachedTop =915
                    LayoutCachedWidth =15344
                    LayoutCachedHeight =1215
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3030
                    Top =60
                    Width =945
                    Height =255
                    FontSize =8
                    TabIndex =3
                    Name ="LeasePaidAdvArr"
                    GroupTable =1
                    BottomPadding =45

                    LayoutCachedLeft =3030
                    LayoutCachedTop =60
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
                    Left =18480
                    Width =300
                    Height =270
                    FontSize =10
                    TabIndex =16
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

                    LayoutCachedLeft =18480
                    LayoutCachedWidth =18780
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
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
                    Left =4035
                    Width =0
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="OptionTypeID"
                    ControlSource ="OptionTypeID"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4035
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9300
                    Width =1080
                    Height =270
                    FontSize =9
                    TabIndex =9
                    Name ="EndOrActualDate"
                    ControlSource ="EndOrActualDate"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Accepted Date of the NEW Unit"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9300
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
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
                    Width =960
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
                    LayoutCachedWidth =990
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
                    TabIndex =19
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
                    Left =8160
                    Width =1080
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="StartOrPurposedDate"
                    ControlSource ="StartOrPurposedDate"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Estimated Delivery Date of the NEW Unit"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8160
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
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
                    Left =4095
                    Width =15
                    Height =270
                    FontSize =10
                    TabIndex =5
                    BackColor =16317438
                    Name ="ReplacedByScheduleID"
                    ControlSource ="ReplacedByScheduleID"
                    Format ="Standard"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4095
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
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
                    Left =10410
                    Width =4890
                    Height =270
                    FontSize =9
                    TabIndex =10
                    Name ="Notes"
                    ControlSource ="Notes"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="notes..."
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10410
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =9792
                    Left =6705
                    Width =1395
                    Height =270
                    FontSize =10
                    TabIndex =7
                    Name ="ReplacedByUnitID"
                    ControlSource ="ReplacedByUnitID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.UnitID AS AssetID, vw_SixKeys.Unitnum AS [Unit#], [Myear] & \""
                        " \" & [Make] & \" \" & [Model] & \" \" & [Type] AS [Desc], vw_SixKeys.SchID, vw_"
                        "SixKeys.Vin AS VIN FROM vw_SixKeys LEFT JOIN UnitExitOption ON vw_SixKeys.UnitID"
                        " = UnitExitOption.ReplacedByUnitID WHERE (((vw_SixKeys.SGrpID)=[Forms]![EXITOpti"
                        "ons]![ExitOptions_subform_Exchange].[Form]![ReplacedByGroupID]) AND ((UnitExitOp"
                        "tion.ReplacedByUnitID) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) ORDER BY vw_Si"
                        "xKeys.Unitnum;"
                    ColumnWidths ="1008;1296;3888;0;2592"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="This is the exact unit the client chose to replace this unit with, Its Delivery "
                        "and ISD dates will be unsed"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6705
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
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
                    ListWidth =7200
                    Left =4140
                    Width =2520
                    Height =270
                    FontSize =10
                    TabIndex =6
                    BackColor =16317438
                    Name ="ReplacedByGroupID"
                    ControlSource ="ReplacedByGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.SGrpID, [ClientShortName] & \" | \" & [Schedule] & \" | \" & ["
                        "UnitGroup] AS [Schedule-Group], Count(vw_SixKeys.UnitID) AS Units, vw_SixKeys.Sc"
                        "hID FROM vw_SixKeys LEFT JOIN UnitExitOption ON vw_SixKeys.UnitID = UnitExitOpti"
                        "on.ReplacedByUnitID WHERE (((vw_SixKeys.Status)<>\"T\") AND ((vw_SixKeys.MLOrig)"
                        "=\"FA\") AND ((vw_SixKeys.ClientGroupId)=[Forms]![EXITOptions]![ClientGroupID]))"
                        " GROUP BY vw_SixKeys.SGrpID, [ClientShortName] & \" | \" & [Schedule] & \" | \" "
                        "& [UnitGroup], vw_SixKeys.SchID ORDER BY [ClientShortName] & \" | \" & [Schedule"
                        "] & \" | \" & [UnitGroup];"
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

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
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
                    Left =16335
                    Width =1110
                    Height =270
                    FontSize =8
                    TabIndex =14
                    Name ="UserName"
                    ControlSource ="UserName"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16335
                    LayoutCachedWidth =17445
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =17445
                    Width =1035
                    Height =270
                    FontSize =8
                    TabIndex =15
                    Name ="DatePosted"
                    ControlSource ="DatePosted"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17445
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =16
                    ColumnEnd =16
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
                    Left =1050
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

                    LayoutCachedLeft =1050
                    LayoutCachedWidth =1965
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =15840
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =13
                    Name ="ModifyTERM"
                    Caption ="TERM"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Terminate Option for this unit"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15840
                    LayoutCachedWidth =16305
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =790406
                    BorderColor =15527148
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =790406
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15345
                    Width =465
                    Height =270
                    FontSize =8
                    TabIndex =12
                    Name ="ModifyEXTD"
                    Caption ="EXTD"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Extend Option for this unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15345
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =0
                    BorderColor =15527148
                    HoverColor =9211020
                    PressedColor =10319446
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18780
                    Width =2115
                    Height =270
                    FontSize =9
                    TabIndex =17
                    Name ="EarlyGapOverLapMsgTx"
                    ControlSource ="EarlyGapOverLapMsg"
                    ControlTipText ="Expiry and or Overlap messages will appear here, as needed."
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18780
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2025
                    Width =945
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="FinalRentDate"
                    ControlSource ="FinalRentDate"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2025
                    LayoutCachedWidth =2970
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
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20895
                    Width =45
                    Height =270
                    FontSize =9
                    TabIndex =18
                    Name ="UnitPDStartDate"
                    ControlSource ="UnitPDStartDate"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="The Per Diem Start Date of the NEW Unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =20895
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3030
                    Width =945
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="projEol"
                    ControlSource ="=DateAdd(\"d\",-1,DateAdd(\"m\",IIf([LeasePaidAdvArr]=\"ARR\",0,-1),[FinalRentDa"
                        "te]))"
                    ControlTipText ="The Date of the Original Unit's LAST Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3030
                    LayoutCachedWidth =3975
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
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15330
                    Width =14
                    Height =270
                    FontSize =8
                    TabIndex =11
                    Name ="LeaseOptionID"
                    ControlSource ="LeaseOptionID"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15330
                    LayoutCachedWidth =15344
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
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
' See "ExitOptions_subform_Exchange.cls"
