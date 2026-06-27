Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11203
    DatasheetFontHeight =11
    ItemSuffix =66
    Left =8730
    Top =3570
    Right =21360
    Bottom =16530
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x01db86372d6ae440
    End
    OnClose ="[Event Procedure]"
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
        Begin FormHeader
            Height =735
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =540
                    Width =11203
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =11203
                    LayoutCachedHeight =735
                    TabIndex =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Width =4080
                    Height =690
                    FontSize =24
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="SUF Selection"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =690
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =9480
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
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

                    LayoutCachedLeft =9480
                    LayoutCachedTop =180
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =408
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
                    Left =10051
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10051
                    LayoutCachedTop =180
                    LayoutCachedWidth =10801
                    LayoutCachedHeight =408
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
            Height =12240
            Name ="Detail"
            AlternateBackColor =15527148
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =90
                    Top =3608
                    Width =3825
                    Height =2550
                    FontSize =10
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationID , Locationsn AS Location, LState AS State, Count(Unitvin) AS U"
                        "nits FROM vw_SUFReportQuery WHERE [clientID]=[Forms]![SUF_Select_Form]![Company_"
                        "List] GROUP BY LocationID, Locationsn, LState ORDER BY Locationsn, LState;"
                    ColumnWidths ="0;1872;720;720"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =3608
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =6158
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =161
                    Left =6390
                    Top =345
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="line2"
                    Caption ="──┐"
                    FontName ="Arial"
                    LayoutCachedLeft =6390
                    LayoutCachedTop =345
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =4
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =161
                    Left =600
                    Top =3375
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="Label53"
                    Caption ="┌──"
                    FontName ="Arial"
                    LayoutCachedLeft =600
                    LayoutCachedTop =3375
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =3630
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =75
                    Top =600
                    Width =3825
                    Height =2520
                    FontSize =10
                    TabIndex =1
                    Name ="Company_List"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SUFReportQuery.clientID, vw_SUFReportQuery.ClientShortName, Count(vw_S"
                        "UFReportQuery.unitvin) AS Units FROM vw_SUFReportQuery GROUP BY vw_SUFReportQuer"
                        "y.clientID, vw_SUFReportQuery.ClientShortName HAVING (((vw_SUFReportQuery.[Clien"
                        "tShortName]) Like '**')) ORDER BY vw_SUFReportQuery.ClientShortName;"
                    ColumnWidths ="0;2880;360"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =600
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =75
                    Top =270
                    Width =3825
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="Search_Client"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =270
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =510
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =75
                            Top =60
                            Width =1350
                            Height =225
                            FontSize =9
                            FontWeight =700
                            Name ="Label45"
                            Caption ="Search Client:"
                            LayoutCachedLeft =75
                            LayoutCachedTop =60
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =285
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =4200
                    Top =600
                    Width =6768
                    Height =11513
                    FontSize =10
                    TabIndex =3
                    BoundColumn =2
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT clientID, UnitDesc, Unitvin AS VIN, UnitID, InspDate, ParkDate \015\012FR"
                        "OM vw_SUFReportQuery WHERE SchID=0 \015\012GROUP BY clientID, UnitDesc, Unitvin,"
                        " UnitID, InspDate, ParkDate \015\012ORDER BY Unitvin;"
                    ColumnWidths ="0;4176;2592;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4200
                    LayoutCachedTop =600
                    LayoutCachedWidth =10968
                    LayoutCachedHeight =12113
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7905
                    Top =270
                    Width =450
                    Height =255
                    FontSize =9
                    TabIndex =4
                    ForeColor =-2147483615
                    Name ="UnitListcount"
                    ControlSource ="=[UnitsList].ListCount-1"

                    LayoutCachedLeft =7905
                    LayoutCachedTop =270
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =525
                    BorderThemeColorIndex =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =6975
                            Top =270
                            Width =930
                            Height =255
                            FontSize =9
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label49"
                            Caption ="Units In List"
                            LayoutCachedLeft =6975
                            LayoutCachedTop =270
                            LayoutCachedWidth =7905
                            LayoutCachedHeight =525
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =1020
                    Top =3300
                    Width =2895
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="SUFbyLocation"
                    Caption ="Open SUF ALL units in Location"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =3570
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =4200
                    Top =240
                    Width =2205
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="SUFbyUnit"
                    Caption ="Open SUF for Selected Unit"
                    OnClick ="[Event Procedure]"
                    LeftPadding =135
                    TopPadding =105
                    RightPadding =150
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4200
                    LayoutCachedTop =240
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =2
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =4200
                    Top =645
                    Width =3975
                    Height =345
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label55"
                    Caption ="◄ First Select A Client"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =645
                    LayoutCachedWidth =8175
                    LayoutCachedHeight =990
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =4200
                    Top =3690
                    Width =3045
                    Height =360
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label56"
                    Caption ="◄Then Select A Location"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =3690
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =4050
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8418
                    Top =270
                    Width =2550
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =8418
                    LayoutCachedTop =270
                    LayoutCachedWidth =10968
                    LayoutCachedHeight =510
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8418
                            Top =60
                            Width =855
                            Height =225
                            FontSize =9
                            FontWeight =700
                            Name ="Label58"
                            Caption ="Search Vin:"
                            LayoutCachedLeft =8418
                            LayoutCachedTop =60
                            LayoutCachedWidth =9273
                            LayoutCachedHeight =285
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =90
                    Top =6586
                    Width =3825
                    Height =2550
                    FontSize =10
                    TabIndex =8
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SUFReportQuery.SchId, vw_SUFReportQuery.SchNo, Count(vw_SUFReportQuery"
                        ".UnitID) AS Units FROM vw_SUFReportQuery WHERE (((vw_SUFReportQuery.clientID)=[F"
                        "orms]![SUF_Select_Form]![Company_List]) AND ((vw_SUFReportQuery.[SUF Sent]) Is N"
                        "ull)) GROUP BY vw_SUFReportQuery.SchId, vw_SUFReportQuery.SchNo ORDER BY vw_SUFR"
                        "eportQuery.SchId;"
                    ColumnWidths ="0;1872;720;720"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =6586
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =9136
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    Left =600
                    Top =6353
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="Label60"
                    Caption ="┌──"
                    FontName ="Arial"
                    LayoutCachedLeft =600
                    LayoutCachedTop =6353
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =6608
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =1020
                    Top =6278
                    Width =2895
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    Name ="SUFbySchedule"
                    Caption ="Open SUF ALL units in Schedule"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =6278
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =6548
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =4200
                    Top =6900
                    Width =4740
                    Height =345
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label62"
                    Caption ="◄ or Select A Schedule"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =6900
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =7245
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =90
                    Top =9564
                    Width =3825
                    Height =2550
                    FontSize =10
                    TabIndex =10
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SUFReportQuery.SGrpID, vw_SUFReportQuery.SgrpNo AS Grp, Count(vw_SUFRe"
                        "portQuery.UnitID) AS Units FROM vw_SUFReportQuery WHERE (((vw_SUFReportQuery.Sch"
                        "ID)=[Forms]![SUF_Select_Form]![ScheduleList]) AND ((vw_SUFReportQuery.[SUF Sent]"
                        ") Is Null)) GROUP BY vw_SUFReportQuery.SGrpID, vw_SUFReportQuery.SgrpNo ORDER BY"
                        " vw_SUFReportQuery.SGrpID;"
                    ColumnWidths ="0;1872;720;720"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =9564
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =12114
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    Left =600
                    Top =9331
                    Width =480
                    Height =255
                    FontSize =10
                    Name ="Label64"
                    Caption ="┌──"
                    FontName ="Arial"
                    LayoutCachedLeft =600
                    LayoutCachedTop =9331
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =9586
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =1020
                    Top =9256
                    Width =2895
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    Name ="SUFbyGroup"
                    Caption ="Open SUF ALL units in Group"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =9256
                    LayoutCachedWidth =3915
                    LayoutCachedHeight =9526
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
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
' See "SUF_Select_Form.cls"
