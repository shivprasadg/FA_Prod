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
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14688
    DatasheetFontHeight =11
    ItemSuffix =97
    Left =4845
    Top =2715
    Right =21795
    Bottom =17340
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x39278026d95ee440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OnCloseEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Close"
            Argument ="2"
            Argument ="bosubform"
            Argument ="0"
        End
    End
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
            Height =540
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =12345
                    Top =60
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =12345
                    LayoutCachedTop =60
                    LayoutCachedWidth =13095
                    LayoutCachedHeight =288
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =11760
                    Top =60
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
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
                    BackStyle =0

                    LayoutCachedLeft =11760
                    LayoutCachedTop =60
                    LayoutCachedWidth =12315
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =360
                    Width =14688
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =360
                    LayoutCachedWidth =14688
                    LayoutCachedHeight =540
                    TabIndex =4
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =7050
                    Height =495
                    FontSize =20
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label50"
                    Caption ="Purchase Order Child Part Manager"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7170
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10500
                    Top =45
                    Width =843
                    Height =246
                    FontSize =8
                    TabIndex =2
                    Name ="GroupID"
                    TopPadding =0

                    LayoutCachedLeft =10500
                    LayoutCachedTop =45
                    LayoutCachedWidth =11343
                    LayoutCachedHeight =291
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8070
                    Top =45
                    Width =2373
                    Height =246
                    FontSize =8
                    TabIndex =3
                    Name ="ActivePO"
                    TopPadding =0

                    LayoutCachedLeft =8070
                    LayoutCachedTop =45
                    LayoutCachedWidth =10443
                    LayoutCachedHeight =291
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =15975
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =210
                    Top =5685
                    Width =7191
                    Height =480
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label58"
                    Caption =" Asset                       | Part Type                       | Make           "
                        "                     | Model                             | Child Part Cost"
                    LayoutCachedLeft =210
                    LayoutCachedTop =5685
                    LayoutCachedWidth =7401
                    LayoutCachedHeight =6165
                    ForeThemeColorIndex =2
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1380
                    Top =135
                    Width =4740
                    Height =270
                    Name ="SelectCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.clientID, Clients.ClientShNm FROM Clients INNER JOIN Mst"
                        "rLease ON Clients.clientID = MstrLease.Client WHERE (((Clients.ClientShNm) Not L"
                        "ike \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.clientType)=\"gu"
                        "arantor\" Or (Clients.clientType) Like \"*target\")) ORDER BY Clients.ClientShNm"
                        ";"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =135
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =255
                            Top =60
                            Width =5895
                            Height =375
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label1"
                            Caption ="Company:"
                            LayoutCachedLeft =255
                            LayoutCachedTop =60
                            LayoutCachedWidth =6150
                            LayoutCachedHeight =435
                            BackThemeColorIndex =2
                            BackTint =60.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1380
                    Top =480
                    Width =4740
                    Height =270
                    TabIndex =2
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig FROM MstrLease WHERE Client=4971 ORDER By MLNo;"
                    ColumnWidths ="0;2160;720"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =480
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextFontCharSet =238
                            TextAlign =1
                            Left =255
                            Top =405
                            Width =5895
                            Height =390
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label3"
                            Caption ="Lease:"
                            LayoutCachedLeft =255
                            LayoutCachedTop =405
                            LayoutCachedWidth =6150
                            LayoutCachedHeight =795
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =210
                    Top =6195
                    Width =7191
                    Height =4830
                    FontSize =10
                    TabIndex =9
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ChildParts.ChildPartID, Units.UnitID, ChildParts.ChildType, ChildParts.Ch"
                        "ildMake, ChildParts.ChildClass, ChildParts.childFinalCost, Units.GroupREF, Child"
                        "Parts.PONumber FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.Uni"
                        "tREF WHERE (((Units.GroupREF)=[Forms]![ChildPartsToGroupPO]![GroupList]) AND ((C"
                        "hildParts.PONumber) Is Null)) ORDER BY ChildParts.ChildType, ChildParts.ChildMak"
                        "e, ChildParts.ChildClass;"
                    ColumnWidths ="0;1152;1440;1440;1440;1152;0"
                    ControlTipText ="Unit to be moved"
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =210
                    LayoutCachedTop =6195
                    LayoutCachedWidth =7401
                    LayoutCachedHeight =11025
                    BackThemeColorIndex =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7635
                    Top =6780
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="MoveUnits"
                    Caption ="► Add Parts ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7635
                    LayoutCachedTop =6780
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =7140
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =26316
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =26316
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7635
                    Top =6180
                    Height =510
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="GetAll"
                    Caption ="◄Select All         Parts"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7635
                    LayoutCachedTop =6180
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =6690
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =26316
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =26316
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7635
                    Top =8205
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="MoveUnitsBack"
                    Caption ="◄Remove◄"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7635
                    LayoutCachedTop =8205
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =8565
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7635
                    Top =7620
                    Height =510
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="GetAllBack"
                    Caption ="        Select All► Parts   "
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7635
                    LayoutCachedTop =7620
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =8130
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =225
                    Top =5325
                    Width =1365
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =16317438
                    Name ="Label27"
                    ControlSource ="=\"Parts: \" & IIf([UnitsList].[ListCount]-1<0,0,[UnitsList].[ListCount]-1)"

                    LayoutCachedLeft =225
                    LayoutCachedTop =5325
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =5625
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7635
                    Top =8655
                    Width =1395
                    Height =315
                    FontWeight =700
                    Name ="Label51"
                    Caption ="◄- move"
                    LayoutCachedLeft =7635
                    LayoutCachedTop =8655
                    LayoutCachedWidth =9030
                    LayoutCachedHeight =8970
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =7635
                    Top =7200
                    Width =1440
                    Height =315
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label52"
                    Caption ="move -►"
                    LayoutCachedLeft =7635
                    LayoutCachedTop =7200
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =7515
                End
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3465
                    Top =1320
                    Width =2640
                    Height =3885
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchGrp.SGrpID, SchGrp.SGrpNO FROM SchGrp WHERE (((SchGrp.SGrpSchID)=[Form"
                        "s]![ChildPartsToGroupPO]![ScheduleList])) ORDER BY SchGrp.SGrpNO;"
                    ColumnWidths ="0;1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3465
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6105
                    LayoutCachedHeight =5205
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =3465
                            Top =1080
                            Width =1110
                            Height =240
                            FontSize =10
                            FontWeight =700
                            BackColor =15858167
                            Name ="Label20"
                            Caption ="Groups"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =3465
                            LayoutCachedTop =1080
                            LayoutCachedWidth =4575
                            LayoutCachedHeight =1320
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =240
                    Top =1320
                    Width =3030
                    Height =3885
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo, Schedule.SCHNo1, Schedule.SchMLRef FROM S"
                        "chedule WHERE (((Schedule.SchMLRef)=[Forms]![ChildPartsToGroupPO]![LeaseList])) "
                        "ORDER BY Schedule.SchID DESC , Schedule.SchNo, Schedule.SCHNo1;"
                    ColumnWidths ="0;2880"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =5205
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =240
                            Top =1080
                            Width =1110
                            Height =240
                            FontSize =10
                            FontWeight =700
                            BackColor =15858167
                            Name ="Label22"
                            Caption ="Schedules"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1080
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =1320
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =216
                    Top =5928
                    Width =1152
                    Height =252
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    Name ="SearchUnitID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Units.UnitID, Units.GroupREF, ChildParts.PONumber FROM Units INNER JOIN C"
                        "hildParts ON Units.UnitID = ChildParts.UnitREF WHERE (((Units.GroupREF)=[Forms]!"
                        "[ChildPartsToGroupPO]![GroupList]) AND ((ChildParts.PONumber) Is Null));"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =216
                    LayoutCachedTop =5928
                    LayoutCachedWidth =1368
                    LayoutCachedHeight =6180
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1380
                    Top =5928
                    Height =252
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    Name ="SearchPartType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ChildParts.ChildType, Units.GroupREF, ChildParts.PONumber FROM Units INNE"
                        "R JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF GROUP BY ChildParts.Child"
                        "Type, Units.GroupREF, ChildParts.PONumber HAVING (((Units.GroupREF)=[Forms]![Chi"
                        "ldPartsToGroupPO]![GroupList]) AND ((ChildParts.PONumber) Is Null)) ORDER BY Chi"
                        "ldParts.ChildType;"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =5928
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =6180
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2832
                    Top =5928
                    Height =252
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    Name ="SearchChildMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ChildParts.ChildMake FROM Units INNER JOIN ChildParts ON Units.UnitID = C"
                        "hildParts.UnitREF WHERE (((Units.GroupREF)=[Forms]![ChildPartsToGroupPO]![GroupL"
                        "ist]) AND ((ChildParts.PONumber) Is Null)) GROUP BY ChildParts.ChildMake ORDER B"
                        "Y ChildParts.ChildMake;"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2832
                    LayoutCachedTop =5928
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =6180
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4284
                    Top =5928
                    Height =252
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    Name ="SearchChildClass"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ChildParts.ChildClass FROM Units INNER JOIN ChildParts ON Units.UnitID = "
                        "ChildParts.UnitREF WHERE (((Units.GroupREF)=[Forms]![ChildPartsToGroupPO]![Group"
                        "List]) AND ((ChildParts.PONumber) Is Null)) GROUP BY ChildParts.ChildClass ORDER"
                        " BY ChildParts.ChildClass;"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4284
                    LayoutCachedTop =5928
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =6180
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10740
                    Top =4185
                    Width =1155
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    ForeColor =6108695
                    Name ="MakePO"
                    Caption ="Unit PO "
                    ControlTipText ="Open Form"
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

                    LayoutCachedLeft =10740
                    LayoutCachedTop =4185
                    LayoutCachedWidth =11895
                    LayoutCachedHeight =4500
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14474460
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =6108695
                    PressedForeColor =6108695
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =11955
                    Top =4185
                    Width =1095
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    ForeColor =6108695
                    Name ="MakePOTracker"
                    Caption ="PO Tracker"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a66 ,
                        0xb17d4ae7b17d4affb17d4affb17d4ae7b17d4a60000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4af3 ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4ae7000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4acf ,
                        0xb17d4affb17d4a8cb17d4a88b17d4affb17d4ac990a4682a90a468ff90a468ff ,
                        0x90a468e790a46860727272ff727272ff727272ff727272ff00000000b17d4a27 ,
                        0xb17d4acfb17d4a1ab17d4a1cb17d4ac9b17d4a1290a468b790a468ff90a468ff ,
                        0x90a468ff90a468e7727272ffffffffffffffffffffffffffffffffffffffff4e ,
                        0xb17d4a48b17d4ae7b17d4ae7b17d4a2490a4689990a468ff90a4688c90a46888 ,
                        0x90a468ff90a468c9727272ffffffffffffffffffffffffffffffffffffffff1b ,
                        0xb17d4ae4b17d4affb17d4affb17d4ae490a4681390a468cf90a4681a90a4681c ,
                        0x90a468c990a46824727272ffffffffffffffffffffffffffffffffffffffff18 ,
                        0xb17d4ae7b17d4affb17d4affb17d4ae70000000090a4684890a468e790a468e7 ,
                        0x90a4684800000000727272ffffffffffffffffffffffffffffffffffffffff78 ,
                        0xb17d4a3fb17d4ac9b17d4ac9b17d4a3f0000000090a468e490a468ff90a468ff ,
                        0x90a468e400000000727272ffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffff75ffffff1effffff18ffffff6fffffff1e90a468e790a468ff90a468ff ,
                        0x90a468e700000000727272ffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffffffffffffffff7e90a4683f90a468c990a468c9 ,
                        0x90a4683f00000000727272ffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffffffffffb3b3b3ffffffff72ffffff1bffffff1b ,
                        0xffffff6c727272ff727272ffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff
                    End

                    LayoutCachedLeft =11955
                    LayoutCachedTop =4185
                    LayoutCachedWidth =13050
                    LayoutCachedHeight =4500
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14474460
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =6108695
                    PressedForeColor =6108695
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =9225
                    Top =6195
                    Width =5286
                    Height =4800
                    FontSize =10
                    TabIndex =16
                    Name ="POChildPartList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ChildParts.ChildPartID, Units.UnitID, ChildParts.ChildType, ChildParts.Ch"
                        "ildMake, ChildParts.ChildClass, ChildParts.childFinalCost, Units.GroupREF, Child"
                        "Parts.PONumber FROM Units INNER JOIN ChildParts ON Units.UnitID = ChildParts.Uni"
                        "tREF WHERE (((Units.GroupREF)=[Forms]![ChildPartsToGroupPO]![GroupList]) AND ((C"
                        "hildParts.PONumber)=[Forms]![ChildPartsToGroupPO]![ActivePO])) ORDER BY ChildPar"
                        "ts.ChildType, ChildParts.ChildMake, ChildParts.ChildClass;"
                    ColumnWidths ="0;1152;1440;1440;1440;1152;0"
                    ControlTipText ="Unit to be moved"
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =9225
                    LayoutCachedTop =6195
                    LayoutCachedWidth =14511
                    LayoutCachedHeight =10995
                    BackThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =9225
                    Top =5685
                    Width =5286
                    Height =480
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label93"
                    Caption =" Asset                       | Part Type                       | Make           "
                        "                     | Model                       "
                    LayoutCachedLeft =9225
                    LayoutCachedTop =5685
                    LayoutCachedWidth =14511
                    LayoutCachedHeight =6165
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =5325
                    Width =1365
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BackColor =16317438
                    Name ="Text94"
                    ControlSource ="=\"Parts: \" & IIf([POChildPartList].[ListCount]-1<0,0,[POChildPartList].[ListCo"
                        "unt]-1)"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =5325
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =5625
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =6870
                    Top =1290
                    Width =5415
                    Height =2430
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    BoundColumn =2
                    Name ="POList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT GroupPurchaseOrders.GroupId, GroupPurchaseOrders.POVendor, GroupPurchaseO"
                        "rders.PONumber, Vendors.VendorSNm, GroupPurchaseOrders.POType, GroupPurchaseOrde"
                        "rs.POSentDate FROM GroupPurchaseOrders INNER JOIN Vendors ON GroupPurchaseOrders"
                        ".POVendor = Vendors.VendorID WHERE (((GroupPurchaseOrders.GroupId)=[Forms]![Chil"
                        "dPartsToGroupPO]![GroupID]) AND ((GroupPurchaseOrders.PONumber) Is Not Null)) OR"
                        "DER BY GroupPurchaseOrders.GroupId, GroupPurchaseOrders.PONumber, GroupPurchaseO"
                        "rders.POType;"
                    ColumnWidths ="0;0;1440;1440;1152;1008"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =6870
                    LayoutCachedTop =1290
                    LayoutCachedWidth =12285
                    LayoutCachedHeight =3720
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6870
                            Top =1050
                            Width =1575
                            Height =270
                            FontSize =10
                            FontWeight =700
                            BackColor =15858167
                            Name ="Label96"
                            Caption ="Purchase Orders"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =6870
                            LayoutCachedTop =1050
                            LayoutCachedWidth =8445
                            LayoutCachedHeight =1320
                            ForeThemeColorIndex =2
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
' See "ChildPartsToGroupPO.cls"
