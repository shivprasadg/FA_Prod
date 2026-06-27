Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8700
    DatasheetFontHeight =11
    ItemSuffix =1090
    Left =8595
    Top =3180
    Right =18090
    Bottom =13695
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x73732b99723de640
    End
    RecordSource ="SELECT Clients.ClientGroupID, Clients.clientID, Schedule.SchID, SchGrp.SGrpID, S"
        "chGrp.UnitsInGroup, Schedule.SchNo, SchGrp.sGrpNo, SchGrp.PerDiemStartsOn, SchGr"
        "p.BLCD, SchGrp.SGRPNoPD, Clients.ClientShNm, SchGrp.SgrpCommInd FROM ((Clients I"
        "NNER JOIN MstrLease ON Clients.clientID = MstrLease.Client) INNER JOIN Schedule "
        "ON MstrLease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = Sch"
        "Grp.SGrpSchID WHERE (((MstrLease.MLOrig)=\"FA\"));"
    OnCurrent ="[Event Procedure]"
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
            FontSize =9
            FontWeight =700
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2280
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =2085
                    Width =8640
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =2085
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =2280
                    TabIndex =12
                    LayoutGroup =3
                    GroupTable =5
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Width =4320
                    Height =480
                    FontSize =18
                    Name ="Label114"
                    Caption ="Mass Mod- PerDiem Types"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3300
                    Top =960
                    Width =3720
                    Height =240
                    FontSize =8
                    TabIndex =5
                    Name ="PerDiemStartsOnSubmitt"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3300
                    LayoutCachedTop =960
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1200
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =780
                    Top =960
                    Width =2460
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="SgrpCommIndSubmitt"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =780
                    LayoutCachedTop =960
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1200
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =6780
                    Width =645
                    Height =105
                    ColumnOrder =2
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =105
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =225
                    Width =645
                    Height =105
                    FontSize =8
                    TabIndex =1
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =225
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =330
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =7770
                    Top =420
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =3
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7770
                    LayoutCachedTop =420
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =648
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
                Begin CommandButton
                    OverlapFlags =223
                    Left =7200
                    Top =420
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =2
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

                    LayoutCachedLeft =7200
                    LayoutCachedTop =420
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =648
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
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =780
                    Top =1890
                    Width =2460
                    Height =240
                    FontSize =11
                    FontWeight =400
                    BackColor =10040879
                    Name ="CommBaselbl"
                    Caption ="BLCD Basis"
                    GroupTable =8
                    LayoutCachedLeft =780
                    LayoutCachedTop =1890
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =2130
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =3300
                    Top =1890
                    Width =3720
                    Height =240
                    FontSize =11
                    FontWeight =400
                    BackColor =10040879
                    Name ="Label4113"
                    Caption ="Per Diem Starts On"
                    GroupTable =8
                    LayoutCachedLeft =3300
                    LayoutCachedTop =1890
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =2130
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =7080
                    Top =1890
                    Width =1140
                    Height =240
                    FontSize =11
                    FontWeight =400
                    BackColor =10040879
                    Name ="Label90"
                    Caption ="No Per Diems"
                    GroupTable =8
                    LayoutCachedLeft =7080
                    LayoutCachedTop =1890
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =2130
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =120
                    Top =1890
                    Width =600
                    Height =240
                    FontSize =11
                    FontWeight =400
                    BackColor =10040879
                    Name ="Label1057"
                    Caption ="Group"
                    GroupTable =8
                    LayoutCachedLeft =120
                    LayoutCachedTop =1890
                    LayoutCachedWidth =720
                    LayoutCachedHeight =2130
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =2160
                    Left =780
                    Top =1260
                    Width =2460
                    Height =570
                    FontSize =10
                    TabIndex =8
                    BackColor =14151142
                    Name ="SgrpCommIndCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =8

                    LayoutCachedLeft =780
                    LayoutCachedTop =1260
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1830
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =3300
                    Top =1260
                    Width =3720
                    Height =570
                    FontSize =10
                    TabIndex =9
                    BackColor =14151142
                    Name ="PerDiemStartsOnCBD"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =8

                    LayoutCachedLeft =3300
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1830
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7590
                    Top =1260
                    Width =630
                    Height =570
                    TabIndex =10
                    Name ="SGRPNoPDCBO"
                    ControlTipText ="Will prompt you for a Date to use"
                    GroupTable =8

                    LayoutCachedLeft =7590
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =1830
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =1260
                    Width =600
                    Height =570
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="Text1063"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =8

                    LayoutCachedLeft =120
                    LayoutCachedTop =1260
                    LayoutCachedWidth =720
                    LayoutCachedHeight =1830
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =7080
                    Top =1260
                    Width =450
                    Height =570
                    Name ="EmptyCell1064"
                    GroupTable =8
                    LayoutCachedLeft =7080
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =1830
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =120
                    Top =960
                    Width =600
                    Name ="EmptyCell1076"
                    GroupTable =8
                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =720
                    LayoutCachedHeight =1200
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7080
                    Top =960
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="SGRPNoPDSubmitt"
                    Caption ="▼Use This"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7080
                    LayoutCachedTop =960
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =1200
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =7080
                    Height =480
                    FontSize =14
                    FontWeight =700
                    TabIndex =11
                    BorderColor =0
                    Name ="Label1089"
                    ControlSource ="=[ClientShNm] & \" Sch: \" & [SchNo] & \"\""

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                End
            End
        End
        Begin Section
            Height =672
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Top =600
                    Width =8640
                    Height =72
                    TabIndex =5
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =600
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =672
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =2160
                    Left =780
                    Top =30
                    Width =2460
                    Height =576
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="SgrpCommInd"
                    ControlSource ="SgrpCommInd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =6

                    LayoutCachedLeft =780
                    LayoutCachedTop =30
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =606
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =3300
                    Top =30
                    Width =3720
                    Height =576
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="PerDiemStartsOn"
                    ControlSource ="PerDiemStartsOn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =6

                    LayoutCachedLeft =3300
                    LayoutCachedTop =30
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =606
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =7590
                    Top =30
                    Width =630
                    Height =576
                    TabIndex =3
                    Name ="SGRPNoPD"
                    ControlSource ="SGRPNoPD"
                    ControlTipText ="Will prompt you for a Date to use"
                    GroupTable =6

                    LayoutCachedLeft =7590
                    LayoutCachedTop =30
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =606
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =7200
                    Top =144
                    Width =645
                    Height =105
                    FontSize =8
                    TabIndex =4
                    Name ="GroupID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =144
                    LayoutCachedWidth =7845
                    LayoutCachedHeight =249
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =30
                    Width =600
                    Height =576
                    FontSize =14
                    FontWeight =700
                    Name ="GroupNo"
                    ControlSource ="sGrpNo"
                    StatusBarText ="blcd source accepted or in-service"
                    GroupTable =6

                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =720
                    LayoutCachedHeight =606
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =7080
                    Top =30
                    Width =450
                    Height =576
                    Name ="EmptyCell1055"
                    GroupTable =6
                    LayoutCachedLeft =7080
                    LayoutCachedTop =30
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =606
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =6
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
' See "MassMod_PerdiemType.cls"
