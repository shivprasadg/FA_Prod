Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    RecordLocks =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =19584
    DatasheetFontHeight =11
    ItemSuffix =4598
    Left =32565
    Top =5715
    Right =-13126
    Bottom =21615
    Filter ="SGrpID = 8385"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xecd82f1a7d9fe540
    End
    RecordSource ="GroupsFormQuery"
    Caption ="GroupForm"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1515
            Name ="FormHeader"
            BackThemeColorIndex =1
            BackShade =85.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =1320
                    Width =19584
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =1320
                    LayoutCachedWidth =19584
                    LayoutCachedHeight =1515
                    TabIndex =18
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =93
                    Left =1560
                    Top =450
                    Width =18000
                    BorderColor =8210719
                    Name ="Line486"
                    HorizontalAnchor =2
                    LayoutCachedLeft =1560
                    LayoutCachedTop =450
                    LayoutCachedWidth =19560
                    LayoutCachedHeight =450
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1710
                    Top =450
                    Width =8430
                    Height =315
                    ColumnOrder =6
                    TabIndex =10
                    BackColor =-2147483609
                    ForeColor =-2147483617
                    Name ="GroupHeader"
                    ControlSource ="=UCase([clientshnm] & \"    Schedule: \" & [schNO] & \"    Group: \" & [sgrpno])"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1710
                    LayoutCachedTop =450
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =765
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8685
                    Width =1200
                    ColumnOrder =7
                    ForeColor =-2147483617
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    StatusBarText ="Short Name for Reporting"

                    LayoutCachedLeft =8685
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12615
                    ColumnOrder =8
                    TabIndex =3
                    ForeColor =-2147483617
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    StatusBarText ="schedule number starts with 1"

                    LayoutCachedLeft =12615
                    LayoutCachedWidth =14055
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15840
                    Top =900
                    Width =3600
                    Height =300
                    ColumnOrder =9
                    TabIndex =4
                    ForeColor =-2147483617
                    Name ="SGrpNO"
                    ControlSource ="SGrpNO"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Protected"

                    LayoutCachedLeft =15840
                    LayoutCachedTop =900
                    LayoutCachedWidth =19440
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9885
                    ColumnOrder =10
                    TabIndex =1
                    ForeColor =-2147483617
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =9885
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11325
                    Width =1290
                    ColumnOrder =11
                    TabIndex =2
                    ForeColor =-2147483617
                    Name ="MLDt"
                    ControlSource ="MLDt"
                    Format ="Short Date"

                    LayoutCachedLeft =11325
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15060
                    Width =630
                    ColumnOrder =12
                    TabIndex =5
                    ForeColor =-2147483617
                    Name ="MLOrig"
                    ControlSource ="MLOrig"

                    LayoutCachedLeft =15060
                    LayoutCachedWidth =15690
                    LayoutCachedHeight =240
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =105
                    Top =60
                    Width =1455
                    Height =1335
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =6
                    Name ="GroupsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchGrp.SGrpID, SchGrp.sgrpno AS GrpNo, SchGrp.UnitsInGroup AS Units, SchG"
                        "rp.SGrpSchID FROM SchGrp WHERE (((SchGrp.SGrpSchID)=[Forms]![GroupForm]![SchID])"
                        ") ORDER BY SchGrp.sgrpno;"
                    ColumnWidths ="0;720;720"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =105
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1395
                End
                Begin Label
                    OverlapFlags =255
                    Left =1770
                    Top =45
                    Width =6765
                    Height =480
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label485"
                    Caption ="VIEWING UNIT GROUPS"
                    LayoutCachedLeft =1770
                    LayoutCachedTop =45
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =525
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13170
                    Top =510
                    Width =885
                    Height =315
                    ColumnOrder =14
                    TabIndex =11
                    ForeColor =-2147483617
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =13170
                    LayoutCachedTop =510
                    LayoutCachedWidth =14055
                    LayoutCachedHeight =825
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14070
                    Top =510
                    Width =900
                    Height =315
                    ColumnOrder =4
                    TabIndex =12
                    ForeColor =-2147483617
                    Name ="GroupID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =14070
                    LayoutCachedTop =510
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =825
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =1020
                    Width =7320
                    Height =293
                    ColumnOrder =5
                    TabIndex =16
                    Name ="SgrpDesc"
                    ControlSource ="SgrpDesc"
                    StatusBarText ="discription for the RAL Attach Letter"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =1313
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1680
                            Top =780
                            Width =4680
                            Height =293
                            FontSize =10
                            Name ="Label418"
                            Caption ="Group Equipment Description for RAL Attachment"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =780
                            LayoutCachedWidth =6360
                            LayoutCachedHeight =1073
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15396
                    Top =516
                    Width =3696
                    ColumnOrder =3
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="UnitGroupID"
                    ControlSource ="=\"SchID: \" & [SchID] & \"  GrpID: \" & [SGrpID] & \" BankID: \" & [AssigneeBan"
                        "kID]"
                    GroupTable =26
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15396
                    LayoutCachedTop =516
                    LayoutCachedWidth =19092
                    LayoutCachedHeight =756
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =26
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =10620
                    Top =240
                    Width =450
                    Height =450
                    FontSize =9
                    TabIndex =9
                    Name ="LockUnLockBtn"
                    Caption ="L"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Click To Un-Lock"
                    Picture ="C:\\FAAS\\Images\\Locked.png"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10620
                    LayoutCachedTop =240
                    LayoutCachedWidth =11070
                    LayoutCachedHeight =690
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderThemeColorIndex =4
                    HoverColor =52479
                    PressedColor =15527148
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =11160
                    Top =345
                    Width =1320
                    Height =240
                    FontSize =9
                    ForeColor =12349952
                    Name ="FormLocked"
                    Caption ="Form Locked"
                    LayoutCachedLeft =11160
                    LayoutCachedTop =345
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =585
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =95
                    PictureType =1
                    Left =18285
                    Top =135
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =18285
                    LayoutCachedTop =135
                    LayoutCachedWidth =19035
                    LayoutCachedHeight =363
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
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
                    OverlapFlags =215
                    Left =17730
                    Top =135
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
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

                    LayoutCachedLeft =17730
                    LayoutCachedTop =135
                    LayoutCachedWidth =18285
                    LayoutCachedHeight =363
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
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
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10455
                    Top =855
                    Width =1980
                    Height =270
                    ColumnOrder =2
                    TabIndex =14
                    ForeColor =-2147483617
                    Name ="Assigned_Indicator"
                    ControlSource ="Assigned_Indicator"

                    LayoutCachedLeft =10455
                    LayoutCachedTop =855
                    LayoutCachedWidth =12435
                    LayoutCachedHeight =1125
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Top =885
                    Width =1980
                    Height =270
                    ColumnOrder =0
                    TabIndex =15
                    ForeColor =-2147483617
                    Name ="Assignee_Bank"
                    ControlSource ="Assignee_Bank"

                    LayoutCachedLeft =12600
                    LayoutCachedTop =885
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =1155
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Width =660
                    TabIndex =17
                    ForeColor =-2147483617
                    Name ="MLID"
                    ControlSource ="MLID"
                    StatusBarText ="schedule number starts with 1"

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =14400
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =4
                    OverlapFlags =93
                    BorderLineStyle =3
                    Left =13380
                    Top =2760
                    Width =1016
                    Height =840
                    Name ="Box4522"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =2760
                    LayoutCachedWidth =14396
                    LayoutCachedHeight =3600
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =4
                    OverlapFlags =223
                    BorderLineStyle =3
                    Left =13380
                    Top =2535
                    Width =1016
                    Name ="Box4503"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =2535
                    LayoutCachedWidth =14396
                    LayoutCachedHeight =3255
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =4
                    OverlapFlags =223
                    BorderLineStyle =3
                    Left =13380
                    Top =2299
                    Width =1016
                    Name ="Box4350"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =2299
                    LayoutCachedWidth =14396
                    LayoutCachedHeight =3019
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =4
                    OverlapFlags =223
                    BorderLineStyle =3
                    Left =13380
                    Top =2539
                    Width =1016
                    Height =240
                    Name ="Box4351"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =2539
                    LayoutCachedWidth =14396
                    LayoutCachedHeight =2779
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =2940
                    Width =1299
                    FontSize =9
                    TabIndex =45
                    BackColor =15858167
                    BorderColor =14270637
                    ForeColor =8210719
                    Name ="CalcSalesRV"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14535
                    LayoutCachedTop =2940
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =3180
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Left =14025
                            Top =2940
                            Width =480
                            Height =240
                            FontSize =10
                            BorderColor =8210719
                            ForeColor =6108695
                            Name ="Label534"
                            Caption ="LO RV"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =14025
                            LayoutCachedTop =2940
                            LayoutCachedWidth =14505
                            LayoutCachedHeight =3180
                            RowStart =2
                            RowEnd =2
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =3060
                    Top =3720
                    Width =12855
                    Height =900
                    Name ="Box3992"
                    LayoutCachedLeft =3060
                    LayoutCachedTop =3720
                    LayoutCachedWidth =15915
                    LayoutCachedHeight =4620
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =60
                    Width =115
                    Height =1140
                    Name ="Box4285"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =175
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =12915
                    Top =1860
                    Width =3045
                    Height =240
                    FontSize =10
                    Name ="Label3959"
                    Caption ="Calculated Group / Unit Roll Up  "
                    LayoutCachedLeft =12915
                    LayoutCachedTop =1860
                    LayoutCachedWidth =15960
                    LayoutCachedHeight =2100
                    BackThemeColorIndex =1
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =16005
                    Top =3210
                    Width =3375
                    Height =240
                    FontSize =10
                    BackColor =26112
                    Name ="Label4326"
                    Caption ="Bank Client Account Info"
                    LayoutCachedLeft =16005
                    LayoutCachedTop =3210
                    LayoutCachedWidth =19380
                    LayoutCachedHeight =3450
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =16260
                    Top =2364
                    Width =1080
                    Height =648
                    FontSize =10
                    BackColor =6737151
                    BorderColor =8355711
                    Name ="Label4325"
                    Caption ="Rent Assignment\015\012Letter"
                    GroupTable =60
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =16260
                    LayoutCachedTop =2364
                    LayoutCachedWidth =17340
                    LayoutCachedHeight =3012
                    RowStart =1
                    RowEnd =2
                    LayoutGroup =13
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =60
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =95
                    Left =10110
                    Width =2805
                    Height =2040
                    Name ="Box3961"
                    LayoutCachedLeft =10110
                    LayoutCachedWidth =12915
                    LayoutCachedHeight =2040
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =135
                    Top =1080
                    Width =12720
                    Height =900
                    Name ="Box4318"
                    LayoutCachedLeft =135
                    LayoutCachedTop =1080
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =1980
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =223
                    Left =180
                    Width =9930
                    Height =1140
                    Name ="Box4283"
                    LayoutCachedLeft =180
                    LayoutCachedWidth =10110
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =95
                    Left =2055
                    Top =2820
                    Width =10860
                    Height =900
                    Name ="Box503"
                    LayoutCachedLeft =2055
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12915
                    LayoutCachedHeight =3720
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =135
                    Top =2040
                    Width =1845
                    Height =2580
                    BackColor =16317438
                    Name ="Box162"
                    LayoutCachedLeft =135
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =4620
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =223
                    Left =15915
                    Width =101
                    Height =4752
                    Name ="Box515"
                    LayoutCachedLeft =15915
                    LayoutCachedWidth =16016
                    LayoutCachedHeight =4752
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =223
                    Left =3255
                    Top =2025
                    Width =9660
                    Height =780
                    Name ="Box4056"
                    LayoutCachedLeft =3255
                    LayoutCachedTop =2025
                    LayoutCachedWidth =12915
                    LayoutCachedHeight =2805
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =255
                    Top =90
                    Width =1620
                    Height =210
                    FontSize =10
                    Name ="grpfaidlbl"
                    Caption ="FAID"
                    LayoutCachedLeft =255
                    LayoutCachedTop =90
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =4065
                    Top =3330
                    Width =600
                    Height =240
                    FontSize =10
                    Name ="Label498"
                    Caption ="└────"
                    LayoutCachedLeft =4065
                    LayoutCachedTop =3330
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =3570
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3525
                    Top =300
                    Width =750
                    Height =330
                    TabIndex =4
                    Name ="UnitsInGroup"
                    ControlSource ="UnitsInGroup"
                    ValidationRule ="Is Not Null And >0"
                    ValidationText ="No of Units is a required field on the group record."
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =3525
                    LayoutCachedTop =300
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5595
                    Top =300
                    Width =1500
                    Height =330
                    TabIndex =6
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Total Cost to be syndicated"
                    ValidationRule =">=0"
                    ValidationText ="Equipment cost  must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Total Cost to be syndicated"

                    LayoutCachedLeft =5595
                    LayoutCachedTop =300
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5595
                    Top =660
                    Width =1500
                    FontSize =10
                    TabIndex =22
                    Name ="CapLeaseBuyOut"
                    ControlSource ="CapLeaseBuyOut"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="buyout price at expiry of capital lease"

                    LayoutCachedLeft =5595
                    LayoutCachedTop =660
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2175
                    Top =1410
                    Width =1320
                    Height =360
                    TabIndex =27
                    Name ="BLCD"
                    ControlSource ="BLCD"
                    Format ="Short Date"
                    StatusBarText ="All units in group must be Accepted Status A to set the BLCD"
                    ControlTipText ="All Units in group must be Accepted Status A to set the BLCD"

                    LayoutCachedLeft =2175
                    LayoutCachedTop =1410
                    LayoutCachedWidth =3495
                    LayoutCachedHeight =1770
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11595
                    Top =1410
                    Width =1260
                    Height =360
                    TabIndex =34
                    Name ="AssignmentDate\015\012"
                    ControlSource ="AssignmentDate"
                    Format ="Short Date"
                    StatusBarText ="bank assignment date"
                    AfterUpdate ="[Event Procedure]"
                    EventProcPrefix ="AssignmentDate__"

                    LayoutCachedLeft =11595
                    LayoutCachedTop =1410
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =1770
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9795
                    Top =2280
                    Width =510
                    ColumnWidth =1365
                    FontSize =10
                    TabIndex =40
                    BackColor =-2147483600
                    ForeColor =16777215
                    Name ="ExchangeITMonth"
                    ControlSource ="ExchangeITMonth"
                    StatusBarText ="Exchange Month Not notice date"
                    ValidationRule =">=0"
                    ValidationText ="Must be >= 0"
                    DefaultValue ="0"
                    ControlTipText ="Exchange Month Not notice date"
                    BottomPadding =150

                    LayoutCachedLeft =9795
                    LayoutCachedTop =2280
                    LayoutCachedWidth =10305
                    LayoutCachedHeight =2520
                    ColumnStart =4
                    ColumnEnd =4
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =8805
                            Top =2280
                            Width =930
                            Height =240
                            FontSize =10
                            Name ="Label14"
                            Caption ="@ Months"
                            BottomPadding =150
                            LayoutCachedLeft =8805
                            LayoutCachedTop =2280
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =2520
                            ColumnStart =3
                            ColumnEnd =3
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3687
                    Top =1410
                    Width =960
                    Height =360
                    ColumnWidth =1455
                    TabIndex =28
                    Name ="LeaseTermPrimary"
                    ControlSource ="LeaseTermPrimary"
                    StatusBarText ="primary term in months"

                    LayoutCachedLeft =3687
                    LayoutCachedTop =1410
                    LayoutCachedWidth =4647
                    LayoutCachedHeight =1770
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =238
                            TextAlign =2
                            Left =3508
                            Top =1200
                            Width =1319
                            Height =210
                            FontSize =10
                            BackColor =14211288
                            Name ="COmmLBL"
                            Caption =" Term Months"
                            LayoutCachedLeft =3508
                            LayoutCachedTop =1200
                            LayoutCachedWidth =4827
                            LayoutCachedHeight =1410
                            BorderThemeColorIndex =1
                            BorderShade =50.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =4
                    OverlapFlags =223
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8698
                    Top =300
                    Width =1320
                    Height =330
                    TabIndex =8
                    Name ="LRFGroup"
                    ControlSource ="LRFGroup"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =8698
                    LayoutCachedTop =300
                    LayoutCachedWidth =10018
                    LayoutCachedHeight =630
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14988
                    Top =540
                    Width =816
                    Height =285
                    FontSize =10
                    TabIndex =12
                    Name ="SGrpFSLPmt"
                    ControlSource ="SGrpFSLPmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    ValidationRule =">=0"
                    ValidationText ="Rent Amounts must be >=0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Tag ="StayOpen"
                    GroupTable =57

                    LayoutCachedLeft =14988
                    LayoutCachedTop =540
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =825
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7202
                    Top =300
                    Width =1410
                    Height =330
                    TabIndex =7
                    Name ="LeasePymtForGroup"
                    ControlSource ="LeasePymtForGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Rent Amounts must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =7202
                    LayoutCachedTop =300
                    LayoutCachedWidth =8612
                    LayoutCachedHeight =630
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =6
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6408
                    Top =2988
                    Width =1056
                    Height =276
                    FontSize =9
                    TabIndex =47
                    BorderColor =10855845
                    Name ="LeaseExtOption1LRF"
                    Format ="Percent"
                    StatusBarText ="1st option lrf"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    GroupTable =13

                    LayoutCachedLeft =6408
                    LayoutCachedTop =2988
                    LayoutCachedWidth =7464
                    LayoutCachedHeight =3264
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =5784
                            Top =2988
                            Width =552
                            Height =276
                            FontSize =10
                            Name ="Label2529"
                            Caption ="LRF%"
                            GroupTable =13
                            LayoutCachedLeft =5784
                            LayoutCachedTop =2988
                            LayoutCachedWidth =6336
                            LayoutCachedHeight =3264
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    Left =3255
                    Top =3150
                    Width =780
                    Height =255
                    FontSize =10
                    FontWeight =700
                    Name ="EOptLBL"
                    Caption ="Options"
                    LayoutCachedLeft =3255
                    LayoutCachedTop =3150
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =3405
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10200
                    Top =510
                    Width =2640
                    Height =360
                    TabIndex =10
                    Name ="AssigneeBankID"
                    ControlSource ="AssigneeBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName FROM Banks ORDER BY Banks.BankShortName"
                        ";"
                    ColumnWidths ="0;2880"
                    StatusBarText ="select from Bank file"
                    OnClick ="[Event Procedure]"
                    Tag ="Protected"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =510
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =870
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =238
                            TextAlign =2
                            Left =10200
                            Top =300
                            Width =2640
                            Height =210
                            FontSize =10
                            Name ="AssLBL"
                            Caption ="Assignee"
                            LayoutCachedLeft =10200
                            LayoutCachedTop =300
                            LayoutCachedWidth =12840
                            LayoutCachedHeight =510
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2040
                    Top =300
                    Width =1365
                    Height =330
                    TabIndex =3
                    Name ="LeaseType"
                    ControlSource ="LeaseType"
                    RowSourceType ="Table/Query"
                    RowSource ="Select LeaseType, LeaseTypeID From LookUp_LeaseTypes"
                    ColumnWidths ="1440;0"
                    Tag ="Protected"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =300
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =2025
                            Top =90
                            Width =1380
                            Height =210
                            FontSize =10
                            Name ="ltypelbl"
                            Caption ="Lease Type"
                            LayoutCachedLeft =2025
                            LayoutCachedTop =90
                            LayoutCachedWidth =3405
                            LayoutCachedHeight =300
                            BackThemeColorIndex =1
                            BackShade =50.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5520
                    Top =2280
                    Width =3226
                    FontSize =10
                    TabIndex =39
                    Name ="ExchangeITType"
                    ControlSource ="ExchangeITType"
                    RowSourceType ="Value List"
                    RowSource ="Specified Mo;Any Time-90 day Notice;\"Quarterly starting at...\""
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =5520
                    LayoutCachedTop =2280
                    LayoutCachedWidth =8746
                    LayoutCachedHeight =2520
                    ColumnStart =2
                    ColumnEnd =2
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4200
                            Top =2280
                            Width =1260
                            Height =240
                            FontSize =10
                            Name ="EXTypLBL"
                            Caption ="Exchange Type:"
                            BottomPadding =150
                            LayoutCachedLeft =4200
                            LayoutCachedTop =2280
                            LayoutCachedWidth =5460
                            LayoutCachedHeight =2520
                            ColumnStart =1
                            ColumnEnd =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18090
                    Top =3720
                    Width =1260
                    FontSize =10
                    FontWeight =700
                    TabIndex =57
                    Name ="SgrpBnkNo"
                    ControlSource ="SgrpBnkNo"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="StayOpen"

                    LayoutCachedLeft =18090
                    LayoutCachedTop =3720
                    LayoutCachedWidth =19350
                    LayoutCachedHeight =3960
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16050
                    Top =3720
                    Width =1995
                    FontSize =10
                    FontWeight =700
                    TabIndex =56
                    Name ="SchBnkClNO"
                    ControlSource ="SchBnkClNO"
                    Tag ="StayOpen"

                    LayoutCachedLeft =16050
                    LayoutCachedTop =3720
                    LayoutCachedWidth =18045
                    LayoutCachedHeight =3960
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =16050
                            Top =3480
                            Width =1995
                            Height =255
                            FontSize =10
                            Name ="Label144"
                            Caption ="Client #:"
                            LayoutCachedLeft =16050
                            LayoutCachedTop =3480
                            LayoutCachedWidth =18045
                            LayoutCachedHeight =3735
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6351
                    Top =1410
                    Width =750
                    Height =360
                    TabIndex =30
                    Name ="SgrpExNot"
                    ControlSource ="LeaseTermNotice"
                    ControlTipText ="Prior notice applies to ExchangeIt and all extension options"

                    LayoutCachedLeft =6351
                    LayoutCachedTop =1410
                    LayoutCachedWidth =7101
                    LayoutCachedHeight =1770
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =360
                    Top =2790
                    Height =255
                    TabIndex =44
                    BorderColor =12632256
                    Name ="SGRPNoPD"
                    ControlSource ="SGRPNoPD"
                    StatusBarText ="no interim rent indicator"

                    LayoutCachedLeft =360
                    LayoutCachedTop =2790
                    LayoutCachedWidth =620
                    LayoutCachedHeight =3045
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =570
                            Top =2790
                            Width =1185
                            Height =255
                            FontSize =10
                            Name ="NOPDLBL"
                            Caption ="No Per Diems"
                            LayoutCachedLeft =570
                            LayoutCachedTop =2790
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =3045
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =360
                    Top =3135
                    Height =270
                    TabIndex =51
                    BorderColor =12632256
                    Name ="SGrpNoSlv"
                    ControlSource ="SGrpNoSlv"
                    StatusBarText ="no sliver rent indicator"

                    LayoutCachedLeft =360
                    LayoutCachedTop =3135
                    LayoutCachedWidth =620
                    LayoutCachedHeight =3405
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =570
                            Top =3105
                            Width =1185
                            Height =270
                            FontSize =10
                            Name ="SLVPLBL"
                            Caption ="Sliver Rent"
                            LayoutCachedLeft =570
                            LayoutCachedTop =3105
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =3375
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =215
                    Top =4680
                    Width =19575
                    Height =9645
                    TabIndex =65
                    Name ="Tabcontrol"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedTop =4680
                    LayoutCachedWidth =19575
                    LayoutCachedHeight =14325
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =60
                            Top =5160
                            Width =19380
                            Height =9030
                            Name ="UNITS"
                            Caption ="View Units"
                            LayoutCachedLeft =60
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =60
                                    Top =5700
                                    Width =19335
                                    Height =7770
                                    Name ="GroupForm_UnitsSub"
                                    SourceObject ="Form.GroupForm_UnitsSub"
                                    LinkChildFields ="SGrpID"
                                    LinkMasterFields ="SGrpID"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =5700
                                    LayoutCachedWidth =19395
                                    LayoutCachedHeight =13470
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =8490
                                    Top =5460
                                    Width =4185
                                    Height =225
                                    FontSize =9
                                    ForeColor =12349952
                                    Name ="Label526"
                                    Caption ="┌────────────────── Dates ────────────────────┐"
                                    LayoutCachedLeft =8490
                                    LayoutCachedTop =5460
                                    LayoutCachedWidth =12675
                                    LayoutCachedHeight =5685
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    ListWidth =6480
                                    Left =4104
                                    Top =5352
                                    Width =3792
                                    Height =276
                                    FontSize =9
                                    TabIndex =1
                                    BackColor =15592953
                                    Name ="CofaSign"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Contacts.contactID, Contacts.contactFullName, Contacts.contactTitle, Cont"
                                        "acts.contactInc, Contacts.contactCofA FROM Contacts WHERE (((Contacts.clientRef)"
                                        "=[Forms]![GroupForm]![clientID]) AND ((Contacts.ContactObsolete)=No));"
                                    ColumnWidths ="0;1728;2880;576;576"
                                    OnClick ="[Event Procedure]"
                                    GroupTable =53

                                    LayoutCachedLeft =4104
                                    LayoutCachedTop =5352
                                    LayoutCachedWidth =7896
                                    LayoutCachedHeight =5628
                                    ColumnStart =1
                                    ColumnEnd =4
                                    LayoutGroup =9
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =53
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =972
                                            Top =5352
                                            Width =3096
                                            Height =276
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label215"
                                            Caption ="Set COA Signer for All Units in Group"
                                            FontName ="Segoe UI"
                                            GroupTable =53
                                            LeftPadding =0
                                            RightPadding =0
                                            LayoutCachedLeft =972
                                            LayoutCachedTop =5352
                                            LayoutCachedWidth =4068
                                            LayoutCachedHeight =5628
                                            LayoutGroup =9
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =53
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="Vendor Info"
                            EventProcPrefix ="Vendor_Info"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =285
                                    Top =5490
                                    Width =16995
                                    Height =570
                                    FontSize =20
                                    Name ="Label710"
                                    Caption =" Vendor Purchase Orders"
                                    FontName ="Segoe UI"
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =285
                                    LayoutCachedTop =5490
                                    LayoutCachedWidth =17280
                                    LayoutCachedHeight =6060
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =0
                                    PictureType =2
                                    Left =285
                                    Top =6045
                                    Width =19008
                                    Height =180
                                    Name ="Image142"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =285
                                    LayoutCachedTop =6045
                                    LayoutCachedWidth =19293
                                    LayoutCachedHeight =6225
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin EmptyCell
                                    Left =372
                                    Top =6180
                                    Width =4452
                                    Name ="EmptyCell670"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =372
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =6420
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4860
                                    Top =6180
                                    Width =744
                                    Height =240
                                    FontSize =9
                                    FontWeight =700
                                    Name ="Label668"
                                    Caption ="Customer"
                                    GroupTable =1
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =6420
                                    ColumnStart =1
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =5640
                                    Top =6180
                                    Width =1320
                                    Name ="EmptyCell672"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =6420
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =7032
                                    Top =6180
                                    Width =1680
                                    Name ="EmptyCell673"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =6420
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =6180
                                    Width =2340
                                    Name ="EmptyCell4400"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =6420
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =6180
                                    Width =1128
                                    Name ="EmptyCell4444"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =6420
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =12384
                                    Top =6180
                                    Width =2568
                                    Name ="EmptyCell674"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =6420
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =15012
                                    Top =6180
                                    Width =1428
                                    Name ="EmptyCell675"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =6420
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =16512
                                    Top =6180
                                    Width =972
                                    Name ="EmptyCell676"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =6420
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =372
                                    Top =6420
                                    Width =4452
                                    Height =285
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label474"
                                    Caption ="Vendor"
                                    FontName ="Segoe UI"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =372
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4860
                                    Top =6420
                                    Width =744
                                    Height =285
                                    FontSize =9
                                    FontWeight =700
                                    Name ="Label582"
                                    Caption ="PO"
                                    GroupTable =1
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =5640
                                    Top =6420
                                    Width =1320
                                    Height =285
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label475"
                                    Caption ="PO Sent Date"
                                    FontName ="Segoe UI"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =7032
                                    Top =6420
                                    Width =1680
                                    Height =285
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label480"
                                    Caption ="PO #"
                                    FontName ="Segoe UI"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =6420
                                    Width =2340
                                    Height =285
                                    Name ="EmptyCell4433"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =6420
                                    Width =1128
                                    Height =285
                                    Name ="EmptyCell4445"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =12384
                                    Top =6420
                                    Width =2568
                                    Height =285
                                    Name ="EmptyCell656"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =15012
                                    Top =6420
                                    Width =1428
                                    Height =285
                                    Name ="EmptyCell657"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =16512
                                    Top =6420
                                    Width =972
                                    Height =285
                                    Name ="EmptyCell658"
                                    GroupTable =1
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =6420
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =6705
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    ListRows =24
                                    ListWidth =5904
                                    Left =372
                                    Top =6744
                                    Width =4452
                                    Height =312
                                    FontSize =10
                                    TabIndex =1
                                    Name ="POVendor"
                                    ControlSource ="POVendor"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                                        " ORDER BY Vendors.VendorName;"
                                    ColumnWidths ="0;2160;1872;2160"
                                    StatusBarText ="link to Vendor Table"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =372
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =4860
                                    Top =6744
                                    Width =255
                                    Height =312
                                    Name ="EmptyCell684"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =5124
                                    Top =6744
                                    Width =480
                                    Height =312
                                    TabIndex =2
                                    Name ="CustPO1"
                                    ControlSource ="CustPO1"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0

                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =5640
                                    Top =6744
                                    Width =1320
                                    Height =312
                                    FontSize =10
                                    TabIndex =3
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="POSentDt"
                                    ControlSource ="POSentDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =255
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7032
                                    Top =6744
                                    Width =1680
                                    Height =312
                                    FontSize =10
                                    TabIndex =4
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PONum1"
                                    ControlSource ="PONum1"
                                    Format ="mm/dd/yy"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =8784
                                    Top =6744
                                    Width =1176
                                    Height =312
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =5
                                    ForeColor =6108695
                                    Name ="MakePO"
                                    Caption ="Unit PO "
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open Form"
                                    GroupTable =1
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

                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =9960
                                    LayoutCachedHeight =7056
                                    PictureCaptionArrangement =5
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =1
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
                                    GroupTable =1
                                    Shadow =-1
                                    QuickStyle =32
                                    QuickStyleMask =-881
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =10032
                                    Top =6744
                                    Width =1092
                                    Height =312
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =6
                                    ForeColor =6108695
                                    Name ="MakePOTracker"
                                    Caption ="PO Tracker"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open Form"
                                    GroupTable =1
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

                                    LayoutCachedLeft =10032
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =7056
                                    PictureCaptionArrangement =5
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =1
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
                                    GroupTable =1
                                    Shadow =-1
                                    QuickStyle =32
                                    QuickStyleMask =-881
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =11196
                                    Top =6744
                                    Width =1128
                                    Height =312
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =7
                                    ForeColor =6108695
                                    Name ="AddParts"
                                    Caption ="Add Parts"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open Form"
                                    GroupTable =1
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

                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =7056
                                    PictureCaptionArrangement =5
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
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
                                    GroupTable =1
                                    Shadow =-1
                                    QuickStyle =32
                                    QuickStyleMask =-881
                                    Overlaps =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12384
                                    Top =6744
                                    Width =2568
                                    Height =312
                                    FontSize =10
                                    TabIndex =8
                                    Name ="Text460"
                                    ControlSource ="=IIf(IsNull([BLCD]),\"Payment Instructions Verified\",Null)"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =15012
                                    Top =6744
                                    Width =1428
                                    Height =312
                                    FontSize =10
                                    TabIndex =9
                                    Name ="Text445"
                                    ControlSource ="=IIf(IsNull([povendor]),Null,IIf(IsNull([BLCD]),DLookUp(\"[VENDORPMTVER]\",\"Ven"
                                        "dors\",\"[vendorid]=forms!groupform![POVendor]\"),Null))"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =16512
                                    Top =6744
                                    Width =972
                                    Height =312
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =10
                                    ForeColor =255
                                    Name ="PO1nt"
                                    ControlSource ="=IIf(IsNull([POSentDt]) Or [mlorig]<>\"FA\" Or [AssignmentDate]>#1/1/2000# Or nz"
                                        "([povendor]) Or (DateDiff(\"d\",DLookUp(\"[VENDORPMTVER]\",\"Vendors\",\"[vendor"
                                        "id]=forms!groupform![PO1Vendor]\"),Date()))<89,Null,\"Re-Verify!\")"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =6744
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =7056
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    ListRows =24
                                    ListWidth =5904
                                    Left =372
                                    Top =7128
                                    Width =4452
                                    Height =312
                                    FontSize =10
                                    TabIndex =11
                                    Name ="PO2Vendor"
                                    ControlSource ="PO2Vendor"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                                        " ORDER BY Vendors.VendorName;"
                                    ColumnWidths ="0;2160;1872;2160"
                                    StatusBarText ="link to Vendor Table"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =372
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =4860
                                    Top =7128
                                    Width =255
                                    Height =312
                                    Name ="EmptyCell683"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =5124
                                    Top =7128
                                    Width =480
                                    Height =312
                                    TabIndex =12
                                    Name ="CustPO2"
                                    ControlSource ="CustPO2"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0

                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =5640
                                    Top =7128
                                    Width =1320
                                    Height =312
                                    FontSize =10
                                    TabIndex =13
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PO2SentDt"
                                    ControlSource ="PO2SentDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7032
                                    Top =7128
                                    Width =1680
                                    Height =312
                                    FontSize =10
                                    TabIndex =14
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PONum2"
                                    ControlSource ="PONum2"
                                    Format ="mm/dd/yy"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =7128
                                    Width =2340
                                    Height =312
                                    Name ="EmptyCell4440"
                                    GroupTable =1
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =7128
                                    Width =1128
                                    Height =312
                                    Name ="EmptyCell4447"
                                    GroupTable =1
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12384
                                    Top =7128
                                    Width =2568
                                    Height =312
                                    FontSize =10
                                    TabIndex =15
                                    Name ="Text462"
                                    ControlSource ="=IIf(IsNull([BLCD]),\"Payment Instructions Verified\",Null)"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =15012
                                    Top =7128
                                    Width =1428
                                    Height =312
                                    FontSize =10
                                    TabIndex =16
                                    Name ="Text451"
                                    ControlSource ="=IIf(IsNull([po2vendor]),Null,IIf(IsNull([BLCD]),DLookUp(\"[VENDORPMTVER]\",\"Ve"
                                        "ndors\",\"[vendorid]=forms!groupform![PO2Vendor]\"),Null))"
                                    FontName ="Segoe UI"
                                    GroupTable =1
                                    TextFormat =1

                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =16512
                                    Top =7128
                                    Width =972
                                    Height =312
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =17
                                    ForeColor =255
                                    Name ="PO2nt"
                                    ControlSource ="=IIf([mlorig]<>\"FA\" Or [AssignmentDate]>#1/1/2000# Or IsNull([PO2SentDt]) Or ("
                                        "DateDiff(\"d\",DLookUp(\"[VENDORPMTVER]\",\"Vendors\",\"[vendorid]=forms!groupfo"
                                        "rm![PO2Vendor]\"),Date()))<89,Null,\"Re-Verify!\")"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =7128
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =7440
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    ListRows =24
                                    ListWidth =5904
                                    Left =372
                                    Top =7512
                                    Width =4452
                                    Height =312
                                    FontSize =10
                                    TabIndex =18
                                    Name ="PO3Vendor"
                                    ControlSource ="PO3Vendor"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                                        " ORDER BY Vendors.VendorName;"
                                    ColumnWidths ="0;2160;1872;2160"
                                    StatusBarText ="link to Vendor Table"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =372
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =4860
                                    Top =7512
                                    Width =255
                                    Height =312
                                    Name ="EmptyCell682"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =5124
                                    Top =7512
                                    Width =480
                                    Height =312
                                    TabIndex =19
                                    Name ="CustPO3"
                                    ControlSource ="CustPO3"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0

                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =5640
                                    Top =7512
                                    Width =1320
                                    Height =312
                                    FontSize =10
                                    TabIndex =20
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PO3SentDt"
                                    ControlSource ="PO3SentDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7032
                                    Top =7512
                                    Width =1680
                                    Height =312
                                    FontSize =10
                                    TabIndex =21
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PONum3"
                                    ControlSource ="PONum3"
                                    Format ="mm/dd/yy"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =7512
                                    Width =2340
                                    Height =312
                                    Name ="EmptyCell4439"
                                    GroupTable =1
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =7512
                                    Width =1128
                                    Height =312
                                    Name ="EmptyCell4448"
                                    GroupTable =1
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12384
                                    Top =7512
                                    Width =2568
                                    Height =312
                                    FontSize =10
                                    TabIndex =22
                                    Name ="Text463"
                                    ControlSource ="=IIf(IsNull([BLCD]),\"Payment Instructions Verified\",Null)"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =15012
                                    Top =7512
                                    Width =1428
                                    Height =312
                                    FontSize =10
                                    TabIndex =23
                                    Name ="Text454"
                                    ControlSource ="=IIf(IsNull([po3vendor]),Null,IIf(IsNull([BLCD]),DLookUp(\"[VENDORPMTVER]\",\"Ve"
                                        "ndors\",\"[vendorid]=forms!groupform![PO3Vendor]\"),Null))"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =16512
                                    Top =7512
                                    Width =972
                                    Height =312
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =24
                                    ForeColor =255
                                    Name ="Po3nt"
                                    ControlSource ="=IIf([mlorig]<>\"FA\" Or [AssignmentDate]>#1/1/2000# Or IsNull([PO3SentDt]) Or ("
                                        "DateDiff(\"d\",DLookUp(\"[VENDORPMTVER]\",\"Vendors\",\"[vendorid]=forms!groupfo"
                                        "rm![PO3Vendor]\"),Date()))<89,Null,\"Re-Verify!\")"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =7512
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =7824
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    ListRows =24
                                    ListWidth =5904
                                    Left =372
                                    Top =7896
                                    Width =4452
                                    Height =312
                                    FontSize =10
                                    TabIndex =25
                                    Name ="PO4Vendor"
                                    ControlSource ="PO4Vendor"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                                        " ORDER BY Vendors.VendorName;"
                                    ColumnWidths ="0;2160;1872;2160"
                                    StatusBarText ="link to Vendor Table"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =372
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =4860
                                    Top =7896
                                    Width =255
                                    Height =312
                                    Name ="EmptyCell681"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =5124
                                    Top =7896
                                    Width =480
                                    Height =312
                                    TabIndex =26
                                    Name ="CustPO4"
                                    ControlSource ="CustPO4"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0

                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =5640
                                    Top =7896
                                    Width =1320
                                    Height =312
                                    FontSize =10
                                    TabIndex =27
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PO4SentDt"
                                    ControlSource ="PO4SentDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7032
                                    Top =7896
                                    Width =1680
                                    Height =312
                                    FontSize =10
                                    TabIndex =28
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="PONum4"
                                    ControlSource ="PONum4"
                                    Format ="mm/dd/yy"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1

                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =7896
                                    Width =2340
                                    Height =312
                                    Name ="EmptyCell4405"
                                    GroupTable =1
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =7896
                                    Width =1128
                                    Height =312
                                    Name ="EmptyCell4449"
                                    GroupTable =1
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12384
                                    Top =7896
                                    Width =2568
                                    Height =312
                                    FontSize =10
                                    TabIndex =29
                                    Name ="Text464"
                                    ControlSource ="=IIf(IsNull([BLCD]),\"Payment Instructions Verified\",Null)"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =15012
                                    Top =7896
                                    Width =1428
                                    Height =312
                                    FontSize =10
                                    TabIndex =30
                                    Name ="Text457"
                                    ControlSource ="=IIf(IsNull([po4vendor]),DLookUp(\"[VENDORPMTVER]\",\"Vendors\",\"[vendorid]=for"
                                        "ms!groupform![PO4Vendor]\"),Null)"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =15012
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =16440
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =9
                                    ColumnEnd =9
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =16512
                                    Top =7896
                                    Width =972
                                    Height =312
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =31
                                    ForeColor =255
                                    Name ="PO4nt"
                                    ControlSource ="=IIf([mlorig]<>\"FA\" Or [AssignmentDate]>#1/1/2000# Or IsNull([PO4SentDt]) Or ("
                                        "DateDiff(\"d\",DLookUp(\"[VENDORPMTVER]\",\"Vendors\",\"[vendorid]=forms!groupfo"
                                        "rm![PO4Vendor]\"),Date()))<89,Null,\"Re-Verify!\")"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =7896
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =8208
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =372
                                    Top =8280
                                    Width =4743
                                    Height =276
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label4408"
                                    Caption ="Enter Unit Quote Numbers >"
                                    FontName ="Segoe UI"
                                    GroupTable =1
                                    RightPadding =0
                                    LayoutCachedLeft =372
                                    LayoutCachedTop =8280
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =8556
                                    RowStart =6
                                    RowEnd =6
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =5124
                                    Top =8280
                                    Width =6000
                                    Height =276
                                    FontSize =10
                                    TabIndex =32
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="Text4398"
                                    ControlSource ="POQuoteNums"
                                    Format ="mm/dd/yy"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =1
                                    LeftPadding =0

                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =8280
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =8556
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =2
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =8280
                                    Width =1128
                                    Height =276
                                    Name ="EmptyCell4450"
                                    GroupTable =1
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =8280
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =8556
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =12384
                                    Top =8280
                                    Width =2568
                                    Height =276
                                    Name ="EmptyCell694"
                                    GroupTable =1
                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =8280
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =8556
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =16512
                                    Top =8280
                                    Width =972
                                    Height =276
                                    TabIndex =33
                                    Name ="PilotRev"
                                    ControlSource ="PilotRev"
                                    StatusBarText ="is a Polot Review to be conducted"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =8280
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =8556
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =15012
                                            Top =8280
                                            Width =1428
                                            Height =276
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label259"
                                            Caption ="Pilot Revised"
                                            FontName ="Segoe UI"
                                            GroupTable =1
                                            LayoutCachedLeft =15012
                                            LayoutCachedTop =8280
                                            LayoutCachedWidth =16440
                                            LayoutCachedHeight =8556
                                            RowStart =6
                                            RowEnd =6
                                            ColumnStart =9
                                            ColumnEnd =9
                                            LayoutGroup =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =1
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =372
                                    Top =8628
                                    Width =4452
                                    Height =276
                                    Name ="EmptyCell701"
                                    GroupTable =1
                                    LayoutCachedLeft =372
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =4824
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =4860
                                    Top =8628
                                    Width =255
                                    Height =276
                                    Name ="EmptyCell702"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =4860
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =5115
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =5124
                                    Top =8628
                                    Width =480
                                    Height =276
                                    Name ="EmptyCell703"
                                    GroupTable =1
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =5124
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =5604
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =5640
                                    Top =8628
                                    Width =1320
                                    Height =276
                                    Name ="EmptyCell704"
                                    GroupTable =1
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =6960
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =7032
                                    Top =8628
                                    Width =1680
                                    Height =276
                                    Name ="EmptyCell705"
                                    GroupTable =1
                                    LayoutCachedLeft =7032
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =8712
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =8784
                                    Top =8628
                                    Width =2340
                                    Height =276
                                    Name ="EmptyCell4407"
                                    GroupTable =1
                                    LayoutCachedLeft =8784
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =11124
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =5
                                    ColumnEnd =6
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =11196
                                    Top =8628
                                    Width =1128
                                    Height =276
                                    Name ="EmptyCell4451"
                                    GroupTable =1
                                    LayoutCachedLeft =11196
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =12324
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =7
                                    ColumnEnd =7
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin EmptyCell
                                    Left =12384
                                    Top =8628
                                    Width =2568
                                    Height =276
                                    Name ="EmptyCell706"
                                    GroupTable =1
                                    LayoutCachedLeft =12384
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =14952
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =8
                                    ColumnEnd =8
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =16512
                                    Top =8628
                                    Width =972
                                    Height =276
                                    FontSize =10
                                    TabIndex =34
                                    Name ="Text256"
                                    ControlSource ="PilotRevdt"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="date Pilot Conducted"
                                    FontName ="Segoe UI"
                                    GroupTable =1

                                    LayoutCachedLeft =16512
                                    LayoutCachedTop =8628
                                    LayoutCachedWidth =17484
                                    LayoutCachedHeight =8904
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =10
                                    ColumnEnd =10
                                    LayoutGroup =1
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =15012
                                            Top =8628
                                            Width =1428
                                            Height =276
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label257"
                                            Caption ="Revised Date"
                                            FontName ="Segoe UI"
                                            GroupTable =1
                                            LayoutCachedLeft =15012
                                            LayoutCachedTop =8628
                                            LayoutCachedWidth =16440
                                            LayoutCachedHeight =8904
                                            RowStart =7
                                            RowEnd =7
                                            ColumnStart =9
                                            ColumnEnd =9
                                            LayoutGroup =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =1
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="Allowances"
                            Caption ="Allowances"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =305
                                    Top =5345
                                    Width =16995
                                    Height =570
                                    FontSize =20
                                    Name ="Label3680"
                                    Caption ="Mileage / Hour Allowances"
                                    FontName ="Segoe UI"
                                    TopPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =305
                                    LayoutCachedTop =5345
                                    LayoutCachedWidth =17300
                                    LayoutCachedHeight =5915
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =0
                                    PictureType =2
                                    Left =305
                                    Top =5900
                                    Width =16995
                                    Height =180
                                    Name ="Image3681"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =305
                                    LayoutCachedTop =5900
                                    LayoutCachedWidth =17300
                                    LayoutCachedHeight =6080
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =411
                                    Top =6125
                                    Width =5325
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3683"
                                    Caption ="Mileage Allowances"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =411
                                    LayoutCachedTop =6125
                                    LayoutCachedWidth =5736
                                    LayoutCachedHeight =6485
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =411
                                    Top =6430
                                    Width =4875
                                    Height =180
                                    Name ="Image3682"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =411
                                    LayoutCachedTop =6430
                                    LayoutCachedWidth =5286
                                    LayoutCachedHeight =6610
                                    TabIndex =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =408
                                    Top =6600
                                    Width =2832
                                    Height =270
                                    FontSize =10
                                    Name ="Label95"
                                    Caption ="Allowance Method"
                                    GroupTable =58
                                    LayoutCachedLeft =408
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =6870
                                    ColumnEnd =1
                                    LayoutGroup =10
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =6600
                                    Width =240
                                    Height =270
                                    Name ="EmptyCell3745"
                                    GroupTable =58
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =6870
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin EmptyCell
                                    Left =3600
                                    Top =6600
                                    Width =372
                                    Height =270
                                    Name ="EmptyCell3746"
                                    GroupTable =58
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =6870
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4044
                                    Top =6600
                                    Width =1320
                                    Height =270
                                    FontSize =9
                                    Name ="Label390"
                                    Caption ="Charged After"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =6870
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =10
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =5436
                                    Top =6600
                                    Width =1188
                                    Height =270
                                    FontSize =9
                                    Name ="Label391"
                                    Caption ="Charged Amt"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =6870
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =10
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =6696
                                    Top =6600
                                    Width =1812
                                    Height =270
                                    FontSize =9
                                    Name ="Label360"
                                    Caption ="Equipment Type"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =6696
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =8508
                                    LayoutCachedHeight =6870
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =10
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1920
                                    Top =6948
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =2
                                    Name ="AllowedMilesPer"
                                    ControlSource ="AllowedMilesPer"
                                    RowSourceType ="Value List"
                                    RowSource ="PER UNIT;IN AGG"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =408
                                            Top =6948
                                            Width =1440
                                            Height =276
                                            FontSize =9
                                            Name ="Label46"
                                            Caption ="Overcharge is"
                                            FontName ="Segoe UI"
                                            GroupTable =58
                                            LayoutCachedLeft =408
                                            LayoutCachedTop =6948
                                            LayoutCachedWidth =1848
                                            LayoutCachedHeight =7224
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =10
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =58
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =6948
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3700"
                                    GroupTable =58
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =6948
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    BackColor =6108695
                                    Name ="Label152"
                                    Caption ="1"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =10
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =6948
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =3
                                    Name ="SGprAllow1"
                                    ControlSource ="SGprAllow1"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =58

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =6948
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =4
                                    Name ="SGrpov1"
                                    ControlSource ="SGrpov1"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =6696
                                    Top =6948
                                    Width =1812
                                    Height =276
                                    FontSize =10
                                    TabIndex =5
                                    Name ="SgrpEq1"
                                    ControlSource ="SgrpEq1"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                                        "lookup_groupTypes.groupType;"
                                    StatusBarText ="For use when group has different equipment to specify over charge for a type"
                                    FontName ="Segoe UI"
                                    ControlTipText ="For use when group has different equipment to specify over charge for a type"
                                    GroupTable =58

                                    LayoutCachedLeft =6696
                                    LayoutCachedTop =6948
                                    LayoutCachedWidth =8508
                                    LayoutCachedHeight =7224
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ListRows =2
                                    Left =1920
                                    Top =7296
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =6
                                    Name ="AllowedMilesPeriod"
                                    ControlSource ="AllowedMilesPeriod"
                                    RowSourceType ="Value List"
                                    RowSource ="P/YEAR;P/TERM"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    AllowValueListEdits =0

                                    ShowOnlyRowSourceValues =255
                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =408
                                            Top =7296
                                            Width =1440
                                            Height =276
                                            FontSize =9
                                            Name ="Label125"
                                            Caption ="Overage Period"
                                            FontName ="Segoe UI"
                                            GroupTable =58
                                            LayoutCachedLeft =408
                                            LayoutCachedTop =7296
                                            LayoutCachedWidth =1848
                                            LayoutCachedHeight =7572
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =10
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =58
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =7296
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3695"
                                    GroupTable =58
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =7296
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    BackColor =6108695
                                    Name ="Label153"
                                    Caption ="2"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =10
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =7296
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =7
                                    Name ="SGprAllow2"
                                    ControlSource ="SGprAllow2"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =58

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =7296
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =8
                                    Name ="SGrpov2"
                                    ControlSource ="SGrpov2"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =6696
                                    Top =7296
                                    Width =1812
                                    Height =276
                                    FontSize =10
                                    TabIndex =9
                                    Name ="SgrpEq2"
                                    ControlSource ="SgrpEq2"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                                        "lookup_groupTypes.groupType;"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =6696
                                    LayoutCachedTop =7296
                                    LayoutCachedWidth =8508
                                    LayoutCachedHeight =7572
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin EmptyCell
                                    Left =408
                                    Top =7644
                                    Height =276
                                    Name ="EmptyCell3703"
                                    GroupTable =58
                                    LayoutCachedLeft =408
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =1848
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin EmptyCell
                                    Left =1920
                                    Top =7644
                                    Width =1320
                                    Height =276
                                    Name ="EmptyCell3704"
                                    GroupTable =58
                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =7644
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3705"
                                    GroupTable =58
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =10
                                    GroupTable =58
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =7644
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    BackColor =6108695
                                    Name ="Label154"
                                    Caption ="3"
                                    FontName ="Segoe UI"
                                    GroupTable =58
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =10
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =7644
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =10
                                    Name ="SGprAllow3"
                                    ControlSource ="SGprAllow3"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =58

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =7644
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =11
                                    Name ="SGrpov3"
                                    ControlSource ="SGrpov3"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =6696
                                    Top =7644
                                    Width =1812
                                    Height =276
                                    FontSize =10
                                    TabIndex =12
                                    Name ="SgrpEq3"
                                    ControlSource ="SgrpEq3"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT lookup_groupTypes.groupType, lookup_groupTypes.grpClass FROM loo"
                                        "kup_groupTypes WHERE (((lookup_groupTypes.grpClass)<>\"MISC\" And (lookup_groupT"
                                        "ypes.grpClass)<>\"OTHER\" And (lookup_groupTypes.grpClass)<>\"BODY\")) ORDER BY "
                                        "lookup_groupTypes.groupType;"
                                    FontName ="Segoe UI"
                                    GroupTable =58

                                    LayoutCachedLeft =6696
                                    LayoutCachedTop =7644
                                    LayoutCachedWidth =8508
                                    LayoutCachedHeight =7920
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =6
                                    ColumnEnd =6
                                    LayoutGroup =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =58
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10251
                                    Top =7265
                                    Width =2280
                                    Height =600
                                    FontSize =9
                                    TabIndex =13
                                    BackColor =13434879
                                    Name ="AllowedMilesNote"
                                    ControlSource ="AllowedMilesNote"
                                    StatusBarText ="special mi provision"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =10251
                                    LayoutCachedTop =7265
                                    LayoutCachedWidth =12531
                                    LayoutCachedHeight =7865
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =8631
                                            Top =7265
                                            Width =1560
                                            Height =600
                                            FontSize =10
                                            Name ="Label385"
                                            Caption ="Other Mileage Provision"
                                            LayoutCachedLeft =8631
                                            LayoutCachedTop =7265
                                            LayoutCachedWidth =10191
                                            LayoutCachedHeight =7865
                                        End
                                    End
                                End
                                Begin Line
                                    OverlapFlags =119
                                    Left =1026
                                    Top =8165
                                    Width =11340
                                    Name ="Line379"
                                    LayoutCachedLeft =1026
                                    LayoutCachedTop =8165
                                    LayoutCachedWidth =12366
                                    LayoutCachedHeight =8165
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =411
                                    Top =8225
                                    Width =5355
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3770"
                                    Caption ="Hours Allowances"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =411
                                    LayoutCachedTop =8225
                                    LayoutCachedWidth =5766
                                    LayoutCachedHeight =8585
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =486
                                    Top =8530
                                    Width =4875
                                    Height =180
                                    Name ="Image3771"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =486
                                    LayoutCachedTop =8530
                                    LayoutCachedWidth =5361
                                    LayoutCachedHeight =8710
                                    TabIndex =14
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =408
                                    Top =8700
                                    Width =2832
                                    Height =270
                                    FontSize =10
                                    Name ="Label3813"
                                    Caption ="Allowance Method"
                                    GroupTable =59
                                    LayoutCachedLeft =408
                                    LayoutCachedTop =8700
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =8970
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =8700
                                    Width =240
                                    Height =270
                                    Name ="EmptyCell3817"
                                    GroupTable =59
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =8700
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =8970
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin EmptyCell
                                    Left =3600
                                    Top =8700
                                    Width =372
                                    Height =270
                                    Name ="EmptyCell3818"
                                    GroupTable =59
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =8700
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =8970
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4044
                                    Top =8700
                                    Width =1320
                                    Height =270
                                    FontSize =9
                                    Name ="Label3821"
                                    Caption ="Charged After"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =8700
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =8970
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =5436
                                    Top =8700
                                    Width =1188
                                    Height =270
                                    FontSize =9
                                    Name ="Label3823"
                                    Caption ="Charged Amt"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =8700
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =8970
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1920
                                    Top =9048
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =15
                                    Name ="AllowedHoursPer"
                                    ControlSource ="AllowedHoursPer"
                                    RowSourceType ="Value List"
                                    RowSource ="PER UNIT;IN AGG"
                                    FontName ="Segoe UI"
                                    GroupTable =59

                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =9048
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =9324
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =408
                                            Top =9048
                                            Width =1440
                                            Height =276
                                            FontSize =9
                                            Name ="Label382"
                                            Caption ="Overcharge is"
                                            FontName ="Segoe UI"
                                            GroupTable =59
                                            LayoutCachedLeft =408
                                            LayoutCachedTop =9048
                                            LayoutCachedWidth =1848
                                            LayoutCachedHeight =9324
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =11
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =59
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =9048
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3786"
                                    GroupTable =59
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =9048
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =9324
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =9048
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    Name ="Label3780"
                                    Caption ="1"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =9048
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =9324
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackShade =50.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =9048
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =16
                                    Name ="SGprAllow4"
                                    ControlSource ="SGprAllow4"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =59

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =9048
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =9324
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =9048
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =17
                                    Name ="SGrpov4"
                                    ControlSource ="SGrpov4"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =59

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =9048
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =9324
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ListRows =2
                                    Left =1920
                                    Top =9396
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =18
                                    Name ="AllowedHoursPeriod"
                                    ControlSource ="AllowedHoursPeriod"
                                    RowSourceType ="Value List"
                                    RowSource ="P/YEAR;P/TERM"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    AllowValueListEdits =0

                                    ShowOnlyRowSourceValues =255
                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =9396
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =9672
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =408
                                            Top =9396
                                            Width =1440
                                            Height =276
                                            FontSize =9
                                            Name ="Label384"
                                            Caption ="Overage Period"
                                            FontName ="Segoe UI"
                                            GroupTable =59
                                            LayoutCachedLeft =408
                                            LayoutCachedTop =9396
                                            LayoutCachedWidth =1848
                                            LayoutCachedHeight =9672
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =11
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =59
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =9396
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3782"
                                    GroupTable =59
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =9396
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =9672
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =9396
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    Name ="Label3809"
                                    Caption ="2"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =9396
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =9672
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackShade =50.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =9396
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =19
                                    Name ="SGprAllow5"
                                    ControlSource ="SGprAllow5"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =59

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =9396
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =9672
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =9396
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =20
                                    Name ="SGrpov5"
                                    ControlSource ="SGrpov5"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =59

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =9396
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =9672
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin EmptyCell
                                    Left =408
                                    Top =9744
                                    Height =276
                                    Name ="EmptyCell3801"
                                    GroupTable =59
                                    LayoutCachedLeft =408
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =1848
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin EmptyCell
                                    Left =1920
                                    Top =9744
                                    Width =1320
                                    Height =276
                                    Name ="EmptyCell3802"
                                    GroupTable =59
                                    LayoutCachedLeft =1920
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =3240
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin EmptyCell
                                    Left =3300
                                    Top =9744
                                    Width =240
                                    Height =276
                                    Name ="EmptyCell3803"
                                    GroupTable =59
                                    LayoutCachedLeft =3300
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =3540
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =59
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3600
                                    Top =9744
                                    Width =372
                                    Height =276
                                    FontSize =9
                                    Name ="Label3810"
                                    Caption ="3"
                                    FontName ="Segoe UI"
                                    GroupTable =59
                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =3972
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =11
                                    BackThemeColorIndex =6
                                    BackShade =50.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4044
                                    Top =9744
                                    Width =1320
                                    Height =276
                                    FontSize =10
                                    TabIndex =21
                                    Name ="SGprAllow6"
                                    ControlSource ="SGprAllow6"
                                    Format ="Standard"
                                    StatusBarText ="Enter all mile overages before hrs"
                                    ValidationRule =">=0"
                                    ValidationText ="Allowances must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Enter all mile overages before hrs"
                                    GroupTable =59

                                    LayoutCachedLeft =4044
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =5364
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5436
                                    Top =9744
                                    Width =1188
                                    Height =276
                                    FontSize =10
                                    TabIndex =22
                                    Name ="SGrpov6"
                                    ControlSource ="SGrpov6"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="charge exceed allowance 1"
                                    ValidationRule =">=0"
                                    ValidationText ="Over use amounts must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    GroupTable =59

                                    LayoutCachedLeft =5436
                                    LayoutCachedTop =9744
                                    LayoutCachedWidth =6624
                                    LayoutCachedHeight =10020
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =11
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =59
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =10251
                                    Top =9365
                                    Width =2280
                                    Height =600
                                    FontSize =9
                                    TabIndex =23
                                    BackColor =13434879
                                    Name ="AllowedHoursNote"
                                    ControlSource ="AllowedHoursNote"
                                    StatusBarText ="special mi provision"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =10251
                                    LayoutCachedTop =9365
                                    LayoutCachedWidth =12531
                                    LayoutCachedHeight =9965
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =8826
                                            Top =9365
                                            Width =1365
                                            Height =465
                                            FontSize =9
                                            Name ="Label388"
                                            Caption ="Other Hours Provision"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =8826
                                            LayoutCachedTop =9365
                                            LayoutCachedWidth =10191
                                            LayoutCachedHeight =9830
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Line
                                    OverlapFlags =119
                                    Left =1026
                                    Top =10175
                                    Width =11340
                                    Name ="Line380"
                                    LayoutCachedLeft =1026
                                    LayoutCachedTop =10175
                                    LayoutCachedWidth =12366
                                    LayoutCachedHeight =10175
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =396
                                    Top =10265
                                    Width =5385
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3825"
                                    Caption ="Warranty Detail"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =396
                                    LayoutCachedTop =10265
                                    LayoutCachedWidth =5781
                                    LayoutCachedHeight =10625
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =411
                                    Top =10570
                                    Width =4875
                                    Height =180
                                    Name ="Image3826"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =411
                                    LayoutCachedTop =10570
                                    LayoutCachedWidth =5286
                                    LayoutCachedHeight =10750
                                    TabIndex =24
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =471
                                    Top =12305
                                    Width =7800
                                    Height =315
                                    FontSize =9
                                    FontWeight =700
                                    Name ="Label419"
                                    Caption ="No. of Units on group adjusted to exclude status O (off lease) for portal "
                                    LayoutCachedLeft =471
                                    LayoutCachedTop =12305
                                    LayoutCachedWidth =8271
                                    LayoutCachedHeight =12620
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="Loan Info"
                            EventProcPrefix ="Loan_Info"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =300
                                    Top =5445
                                    Width =6645
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3847"
                                    Caption ="Loan (Interim) Information"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =300
                                    LayoutCachedTop =5445
                                    LayoutCachedWidth =6945
                                    LayoutCachedHeight =5805
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =330
                                    Top =5750
                                    Width =6735
                                    Height =180
                                    Name ="Image3848"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =330
                                    LayoutCachedTop =5750
                                    LayoutCachedWidth =7065
                                    LayoutCachedHeight =5930
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2316
                                    Top =5928
                                    Width =1632
                                    FontSize =10
                                    TabIndex =1
                                    Name ="SchGrp.SGrpIntBnk"
                                    ControlSource ="SchGrp.SGrpIntBnk"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Clients.ClientShNm FROM Clients WHERE (((Clients.clientType)=\"bank\")) O"
                                        "RDER BY Clients.ClientShNm;"
                                    StatusBarText ="loan for held Portfolio schedule"
                                    FontName ="Segoe UI"
                                    EventProcPrefix ="SchGrp_SGrpIntBnk"
                                    GroupTable =46

                                    LayoutCachedLeft =2316
                                    LayoutCachedTop =5928
                                    LayoutCachedWidth =3948
                                    LayoutCachedHeight =6168
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =456
                                            Top =5928
                                            Width =1788
                                            Height =240
                                            FontSize =9
                                            Name ="Label342"
                                            Caption ="Projected Loan Bank "
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =456
                                            LayoutCachedTop =5928
                                            LayoutCachedWidth =2244
                                            LayoutCachedHeight =6168
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =4008
                                    Top =5928
                                    Width =252
                                    Name ="EmptyCell3877"
                                    GroupTable =46
                                    LayoutCachedLeft =4008
                                    LayoutCachedTop =5928
                                    LayoutCachedWidth =4260
                                    LayoutCachedHeight =6168
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =5
                                    GroupTable =46
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5832
                                    Top =5928
                                    FontSize =10
                                    TabIndex =2
                                    Name ="LoanType\015\012"
                                    ControlSource ="LoanType"
                                    RowSourceType ="Value List"
                                    RowSource ="\"Int\";\"Perm\";\"Loan & Security\""
                                    FontName ="Segoe UI"
                                    EventProcPrefix ="LoanType__"
                                    GroupTable =46

                                    LayoutCachedLeft =5832
                                    LayoutCachedTop =5928
                                    LayoutCachedWidth =7272
                                    LayoutCachedHeight =6168
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4320
                                            Top =5928
                                            Width =1440
                                            Height =240
                                            FontSize =9
                                            Name ="Label114"
                                            Caption ="Loan Type"
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =4320
                                            LayoutCachedTop =5928
                                            LayoutCachedWidth =5760
                                            LayoutCachedHeight =6168
                                            ColumnStart =3
                                            ColumnEnd =3
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2316
                                    Top =6240
                                    Width =1632
                                    FontSize =10
                                    TabIndex =3
                                    Name ="LoanBank\015\012\015\012\015\012"
                                    ControlSource ="LoanBank"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Banks.BankShortName, Banks.BankType FROM Banks WHERE (((Banks.BankType)=\""
                                        "1\")) ORDER BY Banks.BankShortName;"
                                    StatusBarText ="loan for held Portfolio schedule"
                                    FontName ="Segoe UI"
                                    EventProcPrefix ="LoanBank______"
                                    GroupTable =46

                                    LayoutCachedLeft =2316
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =3948
                                    LayoutCachedHeight =6480
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =456
                                            Top =6240
                                            Width =1788
                                            Height =240
                                            FontSize =9
                                            Name ="Label107"
                                            Caption ="Actual Loan Bank "
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =456
                                            LayoutCachedTop =6240
                                            LayoutCachedWidth =2244
                                            LayoutCachedHeight =6480
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =4008
                                    Top =6240
                                    Width =252
                                    Name ="EmptyCell3899"
                                    GroupTable =46
                                    LayoutCachedLeft =4008
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =4260
                                    LayoutCachedHeight =6480
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =5
                                    GroupTable =46
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5832
                                    Top =6240
                                    FontSize =10
                                    TabIndex =4
                                    Name ="LoanDate\015\012"
                                    ControlSource ="LoanDate"
                                    Format ="Short Date"
                                    StatusBarText ="loan date"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    EventProcPrefix ="LoanDate__"
                                    GroupTable =46

                                    LayoutCachedLeft =5832
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =7272
                                    LayoutCachedHeight =6480
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4320
                                            Top =6240
                                            Width =1440
                                            Height =240
                                            FontSize =9
                                            Name ="Label108"
                                            Caption ="Loan Date"
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =4320
                                            LayoutCachedTop =6240
                                            LayoutCachedWidth =5760
                                            LayoutCachedHeight =6480
                                            RowStart =1
                                            RowEnd =1
                                            ColumnStart =3
                                            ColumnEnd =3
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2316
                                    Top =6552
                                    Width =1632
                                    FontSize =10
                                    TabIndex =5
                                    Name ="LoanAmount\015\012"
                                    ControlSource ="LoanAmount"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    EventProcPrefix ="LoanAmount__"
                                    GroupTable =46

                                    LayoutCachedLeft =2316
                                    LayoutCachedTop =6552
                                    LayoutCachedWidth =3948
                                    LayoutCachedHeight =6792
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =456
                                            Top =6552
                                            Width =1788
                                            Height =240
                                            FontSize =9
                                            Name ="Label113"
                                            Caption ="Loan Amount"
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =456
                                            LayoutCachedTop =6552
                                            LayoutCachedWidth =2244
                                            LayoutCachedHeight =6792
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =4008
                                    Top =6552
                                    Width =252
                                    Name ="EmptyCell3878"
                                    GroupTable =46
                                    LayoutCachedLeft =4008
                                    LayoutCachedTop =6552
                                    LayoutCachedWidth =4260
                                    LayoutCachedHeight =6792
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =5
                                    GroupTable =46
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5832
                                    Top =6552
                                    FontSize =10
                                    TabIndex =6
                                    Name ="SgrpPayoffDt"
                                    ControlSource ="SgrpPayoffDt"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="payoff idinividual unit or entire loan"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =46

                                    LayoutCachedLeft =5832
                                    LayoutCachedTop =6552
                                    LayoutCachedWidth =7272
                                    LayoutCachedHeight =6792
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4320
                                            Top =6552
                                            Width =1440
                                            Height =240
                                            FontSize =9
                                            Name ="Label111"
                                            Caption ="Payoff Date"
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =4320
                                            LayoutCachedTop =6552
                                            LayoutCachedWidth =5760
                                            LayoutCachedHeight =6792
                                            RowStart =2
                                            RowEnd =2
                                            ColumnStart =3
                                            ColumnEnd =3
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2316
                                    Top =6864
                                    Width =1632
                                    FontSize =10
                                    TabIndex =7
                                    Name ="SGrpExpDt"
                                    ControlSource ="SGrpExpDt"
                                    Format ="Short Date"
                                    StatusBarText ="date loan required to be paid off"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    GroupTable =46

                                    LayoutCachedLeft =2316
                                    LayoutCachedTop =6864
                                    LayoutCachedWidth =3948
                                    LayoutCachedHeight =7104
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =456
                                            Top =6864
                                            Width =1788
                                            Height =240
                                            FontSize =9
                                            Name ="Label110"
                                            Caption ="Loan (Interim) Expiry"
                                            FontName ="Segoe UI"
                                            GroupTable =46
                                            LayoutCachedLeft =456
                                            LayoutCachedTop =6864
                                            LayoutCachedWidth =2244
                                            LayoutCachedHeight =7104
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =46
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =4008
                                    Top =6864
                                    Width =252
                                    Name ="EmptyCell3937"
                                    GroupTable =46
                                    LayoutCachedLeft =4008
                                    LayoutCachedTop =6864
                                    LayoutCachedWidth =4260
                                    LayoutCachedHeight =7104
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =5
                                    GroupTable =46
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =4320
                                    Top =6864
                                    Width =1440
                                    Height =240
                                    FontSize =9
                                    Name ="Label109"
                                    Caption ="FA Bills"
                                    FontName ="Segoe UI"
                                    GroupTable =46
                                    LayoutCachedLeft =4320
                                    LayoutCachedTop =6864
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =7104
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =5832
                                    Top =6864
                                    Width =1440
                                    TabIndex =8
                                    Name ="SGrpFABill"
                                    ControlSource ="SGrpFABill"
                                    StatusBarText ="if yes FA bills rent if no Lender bills rent"
                                    GroupTable =46

                                    LayoutCachedLeft =5832
                                    LayoutCachedTop =6864
                                    LayoutCachedWidth =7272
                                    LayoutCachedHeight =7104
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =46
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="LO Economics"
                            EventProcPrefix ="LO_Economics"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9026
                            Name ="Economics"
                            Tag ="NoLook"
                            Caption ="LO && Synd Economics"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14186
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =6330
                                    Width =1320
                                    Height =255
                                    FontSize =9
                                    Name ="BLCD_Projected"
                                    ControlSource ="BLCD_Projected"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="The Projected BLCD For This Group"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =6330
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =6585
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =6330
                                            Width =2160
                                            Height =255
                                            FontSize =9
                                            Name ="Label120"
                                            Caption ="BLCD Date"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =6330
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =6585
                                            RowStart =1
                                            RowEnd =1
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =6645
                                    Width =1320
                                    Height =270
                                    FontSize =9
                                    TabIndex =1
                                    Name ="BLCDSetBy"
                                    ControlSource ="BLCDSetBy"
                                    RowSourceType ="Value List"
                                    RowSource ="\"Negotiated\";\"Per Lease\";\"Per ExchangeIT\""
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =6645
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =6915
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =7275
                                    Width =1320
                                    Height =255
                                    FontSize =9
                                    TabIndex =2
                                    Name ="FundDate_Projected"
                                    ControlSource ="FundDate_Projected"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="The Projected SYNDICATED Fund Date"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =7275
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =7530
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =7275
                                            Width =2160
                                            Height =255
                                            FontSize =9
                                            Name ="Label119"
                                            Caption ="Fund Date"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =7275
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =7530
                                            RowStart =4
                                            RowEnd =4
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =7590
                                    Width =1320
                                    Height =255
                                    FontSize =9
                                    TabIndex =3
                                    Name ="Fee_Projected\015\012\015\012\015\012"
                                    ControlSource ="Fee_Projected"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    EventProcPrefix ="Fee_Projected______"
                                    ControlTipText ="The Projected Fee For This Group of Units"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =7590
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =7845
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =7590
                                            Width =2160
                                            Height =255
                                            FontSize =9
                                            Name ="Label121"
                                            Caption ="Syndication Fee"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =7590
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =7845
                                            RowStart =6
                                            RowEnd =6
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =8220
                                    Width =1320
                                    Height =255
                                    FontSize =9
                                    TabIndex =4
                                    Name ="SgrpTermPenalty"
                                    ControlSource ="SgrpTermPenalty"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="Formerly Termination Penalty. This is the amount that will be taken out of the S"
                                        "yndication fee to support Remarketing Sales of units acquired for sale in suppor"
                                        "t of this schedule"
                                    ValidationRule ="<=0"
                                    ValidationText ="Termination Penalty must be<=0 (Remarket Sales Support)"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Formerly Termination Penalty. This is the amount that will be taken out of the S"
                                        "yndication fee to support Remarketing Sales of units acquired for sale in suppor"
                                        "t of this schedule"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =8220
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =8475
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =8220
                                            Width =2160
                                            Height =255
                                            FontSize =9
                                            Name ="Label392"
                                            Caption ="Group Termination Penalty"
                                            FontName ="Segoe UI"
                                            ControlTipText ="Formerly Termination Penalty. This is the amount that will be taken out of the S"
                                                "yndication fee to support Remarketing Sales of units acquired for sale in suppor"
                                                "t of this schedule"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =8220
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =8475
                                            RowStart =7
                                            RowEnd =7
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =8535
                                    Width =1320
                                    FontSize =9
                                    TabIndex =5
                                    Name ="ProjectedNetFee"
                                    ControlSource ="=CCur(Nz([Fee_Projected],0)+Nz([sgrptermPenalty],0))"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Projected Net Fee, After Deducting Penalty (if any)"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =8775
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =8535
                                            Width =2160
                                            Height =240
                                            FontSize =9
                                            Name ="Label412"
                                            Caption ="Net Syndication Fee"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =8535
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =8775
                                            RowStart =8
                                            RowEnd =8
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =8835
                                    Width =1320
                                    Height =285
                                    FontSize =9
                                    TabIndex =6
                                    Name ="PerDiemProjected"
                                    ControlSource ="PerDiemProjected"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    ValidationRule =">=0"
                                    ValidationText ="Per Diem cannot be less than Zero"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Projected MFG Rebate For This Group of Units"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =8835
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =9120
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =8835
                                            Width =2160
                                            Height =285
                                            FontSize =9
                                            Name ="Label1733"
                                            Caption ="Per Diem per Unit"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =8835
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =9120
                                            RowStart =9
                                            RowEnd =9
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2820
                                    Top =9180
                                    Width =1320
                                    Height =285
                                    FontSize =9
                                    TabIndex =7
                                    Name ="LORebate"
                                    ControlSource ="RebateUnitProjected"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    ValidationText ="Rebate must be greater than Zero"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Projected MFG Rebate For This Group of Units"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =9180
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =9465
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =600
                                            Top =9180
                                            Width =2160
                                            Height =285
                                            FontSize =9
                                            Name ="Label159"
                                            Caption ="UTA per Unit"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =600
                                            LayoutCachedTop =9180
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =9465
                                            RowStart =10
                                            RowEnd =10
                                            ColumnEnd =1
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2808
                                    Top =9540
                                    Width =1320
                                    FontSize =9
                                    TabIndex =8
                                    Name ="SalesFSL"
                                    ControlSource ="SalesFSL"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="ATF FSL pmt"
                                    ValidationRule =">=0"
                                    ValidationText ="Sales FSL  Pmt must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Projected Service Fee for These Units"

                                    LayoutCachedLeft =2808
                                    LayoutCachedTop =9540
                                    LayoutCachedWidth =4128
                                    LayoutCachedHeight =9780
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1056
                                            Top =9540
                                            Width =1704
                                            Height =240
                                            FontSize =9
                                            Name ="Label278"
                                            Caption ="FMSS Fee per Unit"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1056
                                            LayoutCachedTop =9540
                                            LayoutCachedWidth =2760
                                            LayoutCachedHeight =9780
                                            RowStart =15
                                            RowEnd =15
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =6285
                                    Width =1065
                                    Height =255
                                    FontSize =9
                                    TabIndex =9
                                    Name ="SGrpSalesIndex"
                                    ControlSource ="SGrpSalesIndex"
                                    RowSourceType ="Value List"
                                    RowSource ="\"LIBOR\";\"SOFR\";\"T-BILL\""
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =6285
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =6540
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =6990
                                            Top =6285
                                            Width =1155
                                            Height =285
                                            FontSize =9
                                            Name ="Label281"
                                            Caption ="Swap Type>"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =6990
                                            LayoutCachedTop =6285
                                            LayoutCachedWidth =8145
                                            LayoutCachedHeight =6570
                                            RowStart =7
                                            RowEnd =7
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =6600
                                    Width =1065
                                    FontSize =9
                                    TabIndex =10
                                    Name ="SGrpSalesIndDt"
                                    ControlSource ="SGrpSalesIndDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =6840
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7245
                                            Top =6600
                                            Width =900
                                            Height =240
                                            FontSize =9
                                            Name ="Label282"
                                            Caption ="Date"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7245
                                            LayoutCachedTop =6600
                                            LayoutCachedWidth =8145
                                            LayoutCachedHeight =6840
                                            RowStart =8
                                            RowEnd =8
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =6900
                                    Width =1065
                                    Height =285
                                    FontSize =9
                                    TabIndex =11
                                    Name ="SGprSalesIndYr"
                                    ControlSource ="SGprSalesIndYr"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =6900
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =7185
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7245
                                            Top =6900
                                            Width =900
                                            Height =285
                                            FontSize =9
                                            Name ="Label283"
                                            Caption ="Term/Year"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7245
                                            LayoutCachedTop =6900
                                            LayoutCachedWidth =8145
                                            LayoutCachedHeight =7185
                                            RowStart =9
                                            RowEnd =9
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =7560
                                    Width =1065
                                    Height =285
                                    FontSize =9
                                    TabIndex =12
                                    Name ="SgrpSalesIndRt"
                                    ControlSource ="SgrpSalesIndRt"
                                    Format ="Percent"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =7560
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =7845
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7140
                                            Top =7560
                                            Width =1035
                                            Height =285
                                            FontSize =9
                                            Name ="Label284"
                                            Caption =" Index Rate"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7140
                                            LayoutCachedTop =7560
                                            LayoutCachedWidth =8175
                                            LayoutCachedHeight =7845
                                            RowStart =10
                                            RowEnd =10
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =7905
                                    Width =1065
                                    Height =285
                                    FontSize =9
                                    TabIndex =13
                                    Name ="SalesSpread"
                                    ControlSource ="=[SgrpSalesYield]-[SgrpSalesIndRt]"
                                    Format ="Percent"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =7905
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =8190
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =6990
                                            Top =7905
                                            Width =1155
                                            Height =285
                                            FontSize =9
                                            Name ="Label286"
                                            Caption ="Spread"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =6990
                                            LayoutCachedTop =7905
                                            LayoutCachedWidth =8145
                                            LayoutCachedHeight =8190
                                            RowStart =11
                                            RowEnd =11
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =7230
                                    Width =1065
                                    Height =285
                                    FontSize =9
                                    TabIndex =14
                                    Name ="SgrpSalesYield"
                                    ControlSource ="SgrpSalesYield"
                                    Format ="Percent"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =7230
                                    LayoutCachedWidth =9285
                                    LayoutCachedHeight =7515
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =6
                                    ColumnEnd =6
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7260
                                            Top =7230
                                            Width =900
                                            Height =285
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label285"
                                            Caption ="MISF Yield"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7260
                                            LayoutCachedTop =7230
                                            LayoutCachedWidth =8160
                                            LayoutCachedHeight =7515
                                            RowStart =12
                                            RowEnd =12
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8220
                                    Top =8895
                                    Width =1080
                                    Height =285
                                    FontSize =9
                                    TabIndex =15
                                    Name ="Bonus"
                                    ControlSource ="SGrpSalesBonus"
                                    RowSourceType ="Value List"
                                    RowSource ="\"NONE\";\"30%\";\"40%\";\"50%\";\"60%\";\"80%\";\"100%\""
                                    StatusBarText ="Bonus Depreciation"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =8895
                                    LayoutCachedWidth =9300
                                    LayoutCachedHeight =9180
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =6330
                                    Width =1425
                                    Height =255
                                    FontSize =9
                                    TabIndex =16
                                    Name ="BLCD_Forecast"
                                    ControlSource ="BLCD_Forecast"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="The BLCD For This Group"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =6330
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =6585
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =6975
                                    Width =1425
                                    FontSize =9
                                    TabIndex =17
                                    Name ="SaleDate"
                                    ControlSource ="SaleDate"
                                    Format ="Short Date"
                                    StatusBarText ="date sold to assignee"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =6975
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =7215
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =7275
                                    Width =1425
                                    Height =255
                                    FontSize =9
                                    TabIndex =18
                                    Name ="FundDate"
                                    ControlSource ="FundDate"
                                    Format ="Short Date"
                                    StatusBarText ="FundDate"
                                    BeforeUpdate ="[Event Procedure]"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =7275
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =7530
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =7890
                                    Width =1425
                                    Height =255
                                    FontSize =9
                                    TabIndex =19
                                    Name ="SalePrice"
                                    ControlSource ="SalePrice"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="sale price to bank"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =7890
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =8145
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =8220
                                    Width =1425
                                    Height =255
                                    FontSize =9
                                    TabIndex =20
                                    Name ="SgrpSynPenalty"
                                    ControlSource ="SgrpSynPenalty"
                                    StatusBarText ="amount/expenses normally reducing the syndication fee - offent money used to sup"
                                        "port Remarketing"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =8220
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =8475
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =8835
                                    Width =1425
                                    Height =285
                                    FontSize =9
                                    TabIndex =21
                                    BackColor =10092543
                                    Name ="PerdiemActual"
                                    ControlSource ="PerDiemActual"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    ValidationRule =">=0"
                                    ValidationText ="Per Diem cannot be less than Zero"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =8835
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =9120
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =9180
                                    Width =1425
                                    Height =285
                                    ColumnWidth =1815
                                    FontSize =9
                                    TabIndex =22
                                    BackColor =10092543
                                    Name ="SynRebate"
                                    ControlSource ="SgrpSynRebate"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    ValidationText ="Rebate must be greater than Zero"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =9180
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =9465
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =9540
                                    Width =1425
                                    FontSize =9
                                    TabIndex =23
                                    Name ="ActualServiceFee"
                                    ControlSource ="=CCur(Nz([SGrpFSLPmt],0)+Nz([SGrpRSLTPmt],0))"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="ATF FSL pmt"
                                    ValidationRule =">=0"
                                    ValidationText ="Sales FSL  Pmt must be >=0"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Projected Service Fee for These Units"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =9540
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =9780
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =6600
                                    Width =1140
                                    FontSize =9
                                    TabIndex =24
                                    Name ="SGrpIndDt"
                                    ControlSource ="SGrpIndDt"
                                    Format ="Short Date"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =6840
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =7
                                    ColumnEnd =7
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =6900
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =25
                                    Name ="SGprIndYr"
                                    ControlSource ="SGprIndYr"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =6900
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =7185
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =7
                                    ColumnEnd =7
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =7560
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =26
                                    Name ="IndexRate"
                                    ControlSource ="IndexRate"
                                    Format ="Percent"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =7560
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =7845
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BackThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =7905
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =27
                                    Name ="Spread"
                                    ControlSource ="=[sgrpyield]-[IndexRate]"
                                    Format ="Percent"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =7905
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =8190
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =7230
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =28
                                    Name ="SgrpYield"
                                    ControlSource ="SgrpYield"
                                    Format ="Percent"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =7230
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =7515
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BackThemeColorIndex =1
                                    BackShade =75.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =8895
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =29
                                    Name ="SGrpBonus"
                                    ControlSource ="SGrpBonus"
                                    RowSourceType ="Value List"
                                    RowSource ="\"NONE\";\"30%\";\"40%\";\"50%\";\"60%\";\"80%\";\"100%\""
                                    StatusBarText ="Bonus Depreciation"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =8895
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =9180
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =7
                                    ColumnEnd =7
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4140
                                    Top =11400
                                    Width =1380
                                    FontSize =9
                                    TabIndex =30
                                    Name ="WithHeldPymts"
                                    ControlSource ="WithHeldPymts"
                                    Format ="Fixed"
                                    ValidationRule =">=0 And <=10"
                                    ValidationText ="Withheld Payments are between 0 and 10"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4140
                                    LayoutCachedTop =11400
                                    LayoutCachedWidth =5520
                                    LayoutCachedHeight =11640
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4140
                                    Top =11700
                                    Width =1380
                                    FontSize =9
                                    TabIndex =31
                                    Name ="WithheldPayments"
                                    ControlSource ="=[LeasePymtForGroup]*[WithHeldPymts]"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="fee for arranging deal on other paper"
                                    ValidationRule =">0 And <10"
                                    ValidationText ="Required 0 or number of pmts bank held back"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =4140
                                    LayoutCachedTop =11700
                                    LayoutCachedWidth =5520
                                    LayoutCachedHeight =11940
                                    RowStart =16
                                    RowEnd =16
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4122
                                    Top =12000
                                    Width =1380
                                    Height =255
                                    FontSize =9
                                    TabIndex =32
                                    Name ="HoldbackResidual"
                                    ControlSource ="HoldbackResidual"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4122
                                    LayoutCachedTop =12000
                                    LayoutCachedWidth =5502
                                    LayoutCachedHeight =12255
                                    RowStart =17
                                    RowEnd =17
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =4122
                                    Top =12315
                                    Width =1380
                                    Height =255
                                    FontSize =9
                                    TabIndex =33
                                    Name ="HoldbackConversion"
                                    ControlSource ="HoldbackConversion"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =4122
                                    LayoutCachedTop =12315
                                    LayoutCachedWidth =5502
                                    LayoutCachedHeight =12570
                                    RowStart =18
                                    RowEnd =18
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4122
                                    Top =12615
                                    Width =1380
                                    Height =255
                                    FontSize =9
                                    TabIndex =34
                                    Name ="NetBankPayment"
                                    ControlSource ="=FormatCurrency([SyndNetFee]+((IIf([SynRebate]>0,[SynRebate],Nz([RebateUnitProje"
                                        "cted],0)))*[UnitsInGroup])-(Nz([StripsAmount],0)+Nz([WithheldPayments],0))-(Nz(["
                                        "HoldbackResidual],0)+Nz([HoldbackConversion],0))+Nz([syndMiscValue1],0)+Nz([synd"
                                        "MiscValue2],0)) & IIf([WithHeldPymts]>0,\" * \",Null)"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =4122
                                    LayoutCachedTop =12615
                                    LayoutCachedWidth =5502
                                    LayoutCachedHeight =12870
                                    RowStart =19
                                    RowEnd =19
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8760
                                    Top =9780
                                    Width =1560
                                    Height =255
                                    FontSize =9
                                    TabIndex =35
                                    Name ="BPChgLO"
                                    ControlSource ="BPChgLO"
                                    StatusBarText ="Apply Adjustment if the Rate Changes By This VALUE. of BP"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =8760
                                    LayoutCachedTop =9780
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =10035
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7440
                                            Top =9780
                                            Width =1200
                                            Height =255
                                            FontSize =9
                                            Name ="Label345"
                                            Caption =" BP Change of "
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7440
                                            LayoutCachedTop =9780
                                            LayoutCachedWidth =8640
                                            LayoutCachedHeight =10035
                                            RowStart =1
                                            RowEnd =1
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =5
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8760
                                    Top =10080
                                    Width =1560
                                    Height =270
                                    FontSize =9
                                    TabIndex =36
                                    Name ="BPChgLO_Up"
                                    ControlSource ="BPChgLO_Up"
                                    Format ="Percent"
                                    StatusBarText ="upadj"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Rate Adjustment - If Upward"

                                    LayoutCachedLeft =8760
                                    LayoutCachedTop =10080
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =10350
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7455
                                            Top =10110
                                            Width =1200
                                            Height =270
                                            FontSize =9
                                            Name ="Label346"
                                            Caption ="Upward"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7455
                                            LayoutCachedTop =10110
                                            LayoutCachedWidth =8655
                                            LayoutCachedHeight =10380
                                            RowStart =2
                                            RowEnd =2
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =5
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8760
                                    Top =10440
                                    Width =1560
                                    FontSize =9
                                    TabIndex =37
                                    Name ="BPChgLO_Down"
                                    ControlSource ="BPChgLO_Down"
                                    Format ="Percent"
                                    StatusBarText ="downadj"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Rate Adjustment - If Downard"

                                    LayoutCachedLeft =8760
                                    LayoutCachedTop =10440
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =10680
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7455
                                            Top =10440
                                            Width =1200
                                            Height =240
                                            FontSize =9
                                            Name ="Label347"
                                            Caption ="Downward"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7455
                                            LayoutCachedTop =10440
                                            LayoutCachedWidth =8655
                                            LayoutCachedHeight =10680
                                            RowStart =3
                                            RowEnd =3
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8820
                                    Top =11445
                                    Width =1260
                                    Height =285
                                    FontSize =9
                                    TabIndex =38
                                    Name ="StripsNum"
                                    ControlSource ="=DateDiff(\"m\",[BLCD],[AssignmentDate])"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8820
                                    LayoutCachedTop =11445
                                    LayoutCachedWidth =10080
                                    LayoutCachedHeight =11730
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8820
                                    Top =11790
                                    Width =1260
                                    Height =285
                                    FontSize =9
                                    TabIndex =39
                                    Name ="StripsAmount"
                                    ControlSource ="=DateDiff(\"m\",[BLCD],[AssignmentDate])*[CalcSalesRent]"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8820
                                    LayoutCachedTop =11790
                                    LayoutCachedWidth =10080
                                    LayoutCachedHeight =12075
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5970
                                    Top =7590
                                    Width =1080
                                    Height =255
                                    FontSize =9
                                    TabIndex =40
                                    Name ="BrokerFee\015\012"
                                    ControlSource ="BrokerFee"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="fee for arranging deal on other paper"
                                    ValidationRule =">0 And <10"
                                    ValidationText ="Required 0 or number of pmts bank held back"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    EventProcPrefix ="BrokerFee__"

                                    LayoutCachedLeft =5970
                                    LayoutCachedTop =7590
                                    LayoutCachedWidth =7050
                                    LayoutCachedHeight =7845
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =4
                                    ColumnEnd =4
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =7590
                                    Width =1425
                                    Height =255
                                    FontSize =9
                                    TabIndex =41
                                    Name ="SynFee"
                                    ControlSource ="SynFee"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =7590
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =7845
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =3
                                    ColumnEnd =3
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =11820
                                    Top =5940
                                    Width =7200
                                    Height =180
                                    Name ="Image3490"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =11820
                                    LayoutCachedTop =5940
                                    LayoutCachedWidth =19020
                                    LayoutCachedHeight =6120
                                    TabIndex =42
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =11880
                                    Top =6600
                                    Width =3180
                                    Height =278
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =43
                                    Name ="FinalEconomicsReportBtn"
                                    Caption ="Final Economics Report"
                                    StatusBarText ="VINS must be entered, Status Must be B and CofA sent date must be nu.."
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Find Next"
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =11880
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =15060
                                    LayoutCachedHeight =6878
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =6108695
                                    BorderColor =6108695
                                    HoverColor =13017476
                                    PressedColor =10040879
                                    HoverForeColor =0
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =32
                                    QuickStyleMask =-885
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =11880
                                    Top =6180
                                    Width =3180
                                    Height =285
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =44
                                    Name ="SyndGroupCheck"
                                    Caption ="Print Synd Group Check"
                                    StatusBarText ="VINS must be entered, Status Must be B and CofA sent date must be nu.."
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Find Next"
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =11880
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =15060
                                    LayoutCachedHeight =6465
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =12750136
                                    BorderColor =6108695
                                    HoverColor =10040879
                                    PressedThemeColorIndex =6
                                    PressedShade =75.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =32
                                    QuickStyleMask =-885
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =11490
                                    Top =12390
                                    Width =3180
                                    Height =285
                                    FontSize =8
                                    TabIndex =45
                                    Name ="UpdtResid"
                                    Caption ="Update Unit Rent, Tax (FSL), Residual"
                                    FontName ="Tahoma"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =11490
                                    LayoutCachedTop =12390
                                    LayoutCachedWidth =14670
                                    LayoutCachedHeight =12675
                                    ColumnStart =1
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    BorderThemeColorIndex =1
                                    BorderShade =50.0
                                    HoverThemeColorIndex =2
                                    HoverTint =40.0
                                    PressedThemeColorIndex =1
                                    PressedShade =85.0
                                    HoverForeThemeColorIndex =0
                                    PressedForeThemeColorIndex =0
                                    QuickStyle =22
                                    QuickStyleMask =-117
                                    Overlaps =1
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =8913
                                    Top =12540
                                    Width =1320
                                    Height =270
                                    FontSize =9
                                    TabIndex =46
                                    Name ="LPAFCost"
                                    ControlSource ="LPAFCost"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="upadj"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Rate Adjustment - If Upward"

                                    LayoutCachedLeft =8913
                                    LayoutCachedTop =12540
                                    LayoutCachedWidth =10233
                                    LayoutCachedHeight =12810
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    CurrencySymbol ="$"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7155
                                            Top =12552
                                            Width =1704
                                            Height =264
                                            FontSize =9
                                            Name ="Label2953"
                                            Caption ="LPAF Eq. Cost"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7155
                                            LayoutCachedTop =12552
                                            LayoutCachedWidth =8859
                                            LayoutCachedHeight =12816
                                            RowStart =7
                                            RowEnd =7
                                            ColumnStart =9
                                            ColumnEnd =9
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =11940
                                    Top =7215
                                    TabIndex =47
                                    Name ="PVFeesYN"
                                    ControlSource ="PVFeesYN"
                                    DefaultValue ="=False"
                                    OnClick ="[Event Procedure]"
                                    Tag ="Acct"

                                    LayoutCachedLeft =11940
                                    LayoutCachedTop =7215
                                    LayoutCachedWidth =12200
                                    LayoutCachedHeight =7455
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =12165
                                            Top =7185
                                            Width =2205
                                            Height =240
                                            FontSize =8
                                            Name ="Label3623"
                                            Caption ="Calculate PV Service Fees @"
                                            FontName ="Segoe UI"
                                            Tag ="Acct"
                                            LayoutCachedLeft =12165
                                            LayoutCachedTop =7185
                                            LayoutCachedWidth =14370
                                            LayoutCachedHeight =7425
                                            ForeThemeColorIndex =1
                                            ForeShade =50.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =11940
                                    Top =7455
                                    TabIndex =48
                                    Name ="ServFeePassTroughYN"
                                    ControlSource ="ServFeePassTroughYN"
                                    DefaultValue ="False"
                                    OnClick ="[Event Procedure]"
                                    Tag ="Acct"

                                    LayoutCachedLeft =11940
                                    LayoutCachedTop =7455
                                    LayoutCachedWidth =12200
                                    LayoutCachedHeight =7695
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =12165
                                            Top =7425
                                            Width =2880
                                            Height =240
                                            FontSize =8
                                            Name ="Label3626"
                                            Caption ="Service Fees are pass through"
                                            FontName ="Segoe UI"
                                            Tag ="Acct"
                                            LayoutCachedLeft =12165
                                            LayoutCachedTop =7425
                                            LayoutCachedWidth =15045
                                            LayoutCachedHeight =7665
                                            ForeThemeColorIndex =1
                                            ForeShade =50.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =14355
                                    Top =7185
                                    Width =675
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =49
                                    ForeColor =9732411
                                    Name ="PVRate"
                                    ControlSource ="PVRate"
                                    Format ="Percent"
                                    DefaultValue ="0.03"
                                    FontName ="Segoe UI"
                                    Tag ="Acct"

                                    LayoutCachedLeft =14355
                                    LayoutCachedTop =7185
                                    LayoutCachedWidth =15030
                                    LayoutCachedHeight =7425
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BorderThemeColorIndex =1
                                    BorderShade =50.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =11940
                                    Top =7695
                                    TabIndex =50
                                    Name ="ShowDateTime"
                                    DefaultValue ="=False"
                                    Tag ="Acct"

                                    LayoutCachedLeft =11940
                                    LayoutCachedTop =7695
                                    LayoutCachedWidth =12200
                                    LayoutCachedHeight =7935
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =12165
                                            Top =7665
                                            Width =2880
                                            Height =240
                                            FontSize =8
                                            Name ="Label3485"
                                            Caption ="Hide 'Date Time' on Report footer"
                                            FontName ="Segoe UI"
                                            Tag ="Acct"
                                            LayoutCachedLeft =12165
                                            LayoutCachedTop =7665
                                            LayoutCachedWidth =15045
                                            LayoutCachedHeight =7905
                                            ForeThemeColorIndex =1
                                            ForeShade =50.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =255
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =14355
                                    Top =6915
                                    Width =675
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =51
                                    Name ="InterimLoanRate"
                                    ControlSource ="InterimLoanRate"
                                    Format ="Percent"
                                    DefaultValue ="0.025"
                                    FontName ="Segoe UI"
                                    Tag ="Acct"

                                    LayoutCachedLeft =14355
                                    LayoutCachedTop =6915
                                    LayoutCachedWidth =15030
                                    LayoutCachedHeight =7155
                                    RowStart =10
                                    RowEnd =10
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BorderThemeColorIndex =1
                                    BorderShade =50.0
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =11940
                                    Top =7935
                                    TabIndex =52
                                    Name ="ShowPreparedBy"
                                    DefaultValue ="=False"
                                    Tag ="Acct"

                                    LayoutCachedLeft =11940
                                    LayoutCachedTop =7935
                                    LayoutCachedWidth =12200
                                    LayoutCachedHeight =8175
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =12165
                                            Top =7905
                                            Width =2880
                                            Height =240
                                            FontSize =8
                                            Name ="Label3488"
                                            Caption ="Hide 'Prepared By' on Report"
                                            FontName ="Segoe UI"
                                            Tag ="Acct"
                                            LayoutCachedLeft =12165
                                            LayoutCachedTop =7905
                                            LayoutCachedWidth =15045
                                            LayoutCachedHeight =8145
                                            ForeThemeColorIndex =1
                                            ForeShade =50.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    DisplayWhen =2
                                    OverlapFlags =247
                                    Left =15495
                                    Top =6180
                                    Width =1590
                                    Height =285
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =53
                                    Name ="ReFreshMe"
                                    Caption ="Re-Calculate All"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run Macro"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =15495
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =17085
                                    LayoutCachedHeight =6465
                                    ForeThemeColorIndex =6
                                    ForeShade =50.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =-2147483633
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    HoverThemeColorIndex =6
                                    HoverShade =75.0
                                    QuickStyle =30
                                    QuickStyleMask =-629
                                    Overlaps =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =7230
                                    Top =5715
                                    Width =3210
                                    Height =180
                                    Name ="Image3627"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =7230
                                    LayoutCachedTop =5715
                                    LayoutCachedWidth =10440
                                    LayoutCachedHeight =5895
                                    TabIndex =54
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =7230
                                    Top =9540
                                    Width =3210
                                    Height =180
                                    Name ="Image3629"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =7230
                                    LayoutCachedTop =9540
                                    LayoutCachedWidth =10440
                                    LayoutCachedHeight =9720
                                    TabIndex =55
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =7110
                                    Top =11235
                                    Width =3210
                                    Height =180
                                    Name ="Image3631"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =7110
                                    LayoutCachedTop =11235
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =11415
                                    TabIndex =56
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =600
                                    Top =5775
                                    Width =4875
                                    Height =180
                                    Name ="Image3633"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =600
                                    LayoutCachedTop =5775
                                    LayoutCachedWidth =5475
                                    LayoutCachedHeight =5955
                                    TabIndex =57
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =540
                                    Top =11205
                                    Width =4875
                                    Height =180
                                    Name ="Image3640"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =540
                                    LayoutCachedTop =11205
                                    LayoutCachedWidth =5415
                                    LayoutCachedHeight =11385
                                    TabIndex =58
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =4350
                                    Top =8535
                                    Width =1275
                                    FontSize =9
                                    TabIndex =59
                                    Name ="SyndNetFee"
                                    ControlSource ="=IIf([SalePrice]>0,([SalePrice]-[OrigEquipCost])+[sgrpsynpenalty],0)"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =4350
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =8775
                                    RowStart =8
                                    RowEnd =8
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2820
                                    Top =5955
                                    Width =1320
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label1183"
                                    Caption ="LO/Target"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =5955
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =6255
                                    ColumnStart =2
                                    ColumnEnd =2
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4185
                                    Top =5955
                                    Width =1395
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    Name ="SyndActualLbl1"
                                    Caption ="Synd Final"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =4185
                                    LayoutCachedTop =5955
                                    LayoutCachedWidth =5580
                                    LayoutCachedHeight =6255
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =600
                                    Top =6645
                                    Width =1050
                                    Height =270
                                    FontSize =9
                                    FontWeight =700
                                    Name ="BLCDHelp"
                                    Caption ="?"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="click for BLCD Set By Help"
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =6645
                                    LayoutCachedWidth =1650
                                    LayoutCachedHeight =6915
                                    RowStart =2
                                    RowEnd =2
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1710
                                    Top =6645
                                    Width =1050
                                    Height =270
                                    FontSize =9
                                    Name ="Label2748"
                                    Caption ="BLCD Set By"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1710
                                    LayoutCachedTop =6645
                                    LayoutCachedWidth =2760
                                    LayoutCachedHeight =6915
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =600
                                    Top =6975
                                    Width =2160
                                    Height =240
                                    FontSize =9
                                    Name ="Label3262"
                                    Caption ="Sale Date"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =6975
                                    LayoutCachedWidth =2760
                                    LayoutCachedHeight =7215
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =5970
                                    Top =7275
                                    Width =960
                                    Height =255
                                    FontSize =9
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22
                                    Name ="Label79"
                                    Caption ="Broker Fee"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    LayoutCachedLeft =5970
                                    LayoutCachedTop =7275
                                    LayoutCachedWidth =6930
                                    LayoutCachedHeight =7530
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =3
                                    ColumnEnd =4
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =600
                                    Top =7890
                                    Width =2160
                                    Height =255
                                    FontSize =9
                                    Name ="Label78"
                                    Caption ="Syndication Sale Price"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =7890
                                    LayoutCachedWidth =2760
                                    LayoutCachedHeight =8145
                                    RowStart =5
                                    RowEnd =5
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =5670
                                    Top =7590
                                    Width =270
                                    Height =255
                                    FontSize =7
                                    Name ="Label80"
                                    Caption ="or"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5670
                                    LayoutCachedTop =7590
                                    LayoutCachedWidth =5940
                                    LayoutCachedHeight =7845
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =11805
                                    Top =5640
                                    Width =3405
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3482"
                                    Caption ="'Economics' Reports"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =11805
                                    LayoutCachedTop =5640
                                    LayoutCachedWidth =15210
                                    LayoutCachedHeight =6000
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =8175
                                    Top =5895
                                    Width =1140
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Label1505"
                                    Caption ="LO/Target"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =8175
                                    LayoutCachedTop =5895
                                    LayoutCachedWidth =9315
                                    LayoutCachedHeight =6195
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =9450
                                    Top =5895
                                    Width =1635
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    Name ="SyndActualLbl2"
                                    Caption ="Synd Final"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =5895
                                    LayoutCachedWidth =11085
                                    LayoutCachedHeight =6195
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =7
                                    ColumnEnd =7
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =7335
                                    Top =11445
                                    Width =1395
                                    Height =285
                                    FontSize =9
                                    Name ="Label139"
                                    Caption ="# of Strips"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =7335
                                    LayoutCachedTop =11445
                                    LayoutCachedWidth =8730
                                    LayoutCachedHeight =11730
                                    RowStart =12
                                    RowEnd =12
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =7335
                                    Top =11790
                                    Width =1395
                                    Height =285
                                    FontSize =9
                                    Name ="SRLabel"
                                    Caption ="Amount of Strips"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =7335
                                    LayoutCachedTop =11790
                                    LayoutCachedWidth =8730
                                    LayoutCachedHeight =12075
                                    RowStart =13
                                    RowEnd =13
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =7260
                                    Top =8895
                                    Width =900
                                    Height =285
                                    FontSize =9
                                    Name ="Label299"
                                    Caption ="Bonus"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =7260
                                    LayoutCachedTop =8895
                                    LayoutCachedWidth =8160
                                    LayoutCachedHeight =9180
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =5
                                    ColumnEnd =5
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1848
                                    Top =11400
                                    Width =2220
                                    Height =240
                                    FontSize =9
                                    Name ="Label427"
                                    Caption ="# Months Advanced Rent*"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1848
                                    LayoutCachedTop =11400
                                    LayoutCachedWidth =4068
                                    LayoutCachedHeight =11640
                                    RowStart =15
                                    RowEnd =15
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1848
                                    Top =11700
                                    Width =2220
                                    Height =240
                                    FontSize =9
                                    Name ="Label2048"
                                    Caption ="Advanced Rent(s) Amount"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1848
                                    LayoutCachedTop =11700
                                    LayoutCachedWidth =4068
                                    LayoutCachedHeight =11940
                                    RowStart =16
                                    RowEnd =16
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1830
                                    Top =12000
                                    Width =2220
                                    Height =255
                                    FontSize =9
                                    Name ="Label2753"
                                    Caption ="Hold-Back Residual"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1830
                                    LayoutCachedTop =12000
                                    LayoutCachedWidth =4050
                                    LayoutCachedHeight =12255
                                    RowStart =17
                                    RowEnd =17
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1830
                                    Top =12315
                                    Width =2220
                                    Height =255
                                    FontSize =9
                                    Name ="Label2758"
                                    Caption ="Hold-Back Conversion"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1830
                                    LayoutCachedTop =12315
                                    LayoutCachedWidth =4050
                                    LayoutCachedHeight =12570
                                    RowStart =18
                                    RowEnd =18
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1830
                                    Top =12615
                                    Width =2220
                                    Height =255
                                    FontSize =9
                                    Name ="Label429"
                                    Caption ="Net Bank Payment"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1830
                                    LayoutCachedTop =12615
                                    LayoutCachedWidth =4050
                                    LayoutCachedHeight =12870
                                    RowStart =19
                                    RowEnd =19
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =1800
                                    Top =12945
                                    Width =2280
                                    Height =240
                                    FontSize =8
                                    Name ="Label2238"
                                    Caption ="* withheld payment(s)"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1800
                                    LayoutCachedTop =12945
                                    LayoutCachedWidth =4080
                                    LayoutCachedHeight =13185
                                    ForeThemeColorIndex =0
                                    ForeTint =50.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =7200
                                    Top =5415
                                    Width =3555
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3628"
                                    Caption ="Index"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =7200
                                    LayoutCachedTop =5415
                                    LayoutCachedWidth =10755
                                    LayoutCachedHeight =5775
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =7230
                                    Top =9240
                                    Width =3585
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3630"
                                    Caption ="Rate Adjustment"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =7230
                                    LayoutCachedTop =9240
                                    LayoutCachedWidth =10815
                                    LayoutCachedHeight =9600
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =7080
                                    Top =10935
                                    Width =3675
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3632"
                                    Caption ="Strips                   "
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =7080
                                    LayoutCachedTop =10935
                                    LayoutCachedWidth =10755
                                    LayoutCachedHeight =11295
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =600
                                    Top =5475
                                    Width =5310
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3634"
                                    Caption ="Economics Detail"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =600
                                    LayoutCachedTop =5475
                                    LayoutCachedWidth =5910
                                    LayoutCachedHeight =5835
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =2910
                                    Top =6975
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    Name ="Label3636"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =2910
                                    LayoutCachedTop =6975
                                    LayoutCachedWidth =4110
                                    LayoutCachedHeight =7215
                                    ForeThemeColorIndex =1
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =2910
                                    Top =7905
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    Name ="Label3637"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =2910
                                    LayoutCachedTop =7905
                                    LayoutCachedWidth =4110
                                    LayoutCachedHeight =8145
                                    ForeThemeColorIndex =1
                                    ForeShade =75.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =540
                                    Top =10905
                                    Width =5310
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3641"
                                    Caption ="Advanced Rent && Hold-Back"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =540
                                    LayoutCachedTop =10905
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =11265
                                    ColumnEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8730
                                    Top =10995
                                    Width =1155
                                    Height =240
                                    FontSize =9
                                    Name ="Label3642"
                                    Caption ="(calculated)"
                                    LayoutCachedLeft =8730
                                    LayoutCachedTop =10995
                                    LayoutCachedWidth =9885
                                    LayoutCachedHeight =11235
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =180
                                    Top =5220
                                    Width =2325
                                    Height =255
                                    FontSize =8
                                    Name ="Label3643"
                                    Caption ="Δ  indicates calculated fields"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =180
                                    LayoutCachedTop =5220
                                    LayoutCachedWidth =2505
                                    LayoutCachedHeight =5475
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5640
                                    Top =7905
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3644"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =7905
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =8145
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10665
                                    Top =7935
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label3645"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10665
                                    LayoutCachedTop =7935
                                    LayoutCachedWidth =10875
                                    LayoutCachedHeight =8175
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10125
                                    Top =11445
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3646"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10125
                                    LayoutCachedTop =11445
                                    LayoutCachedWidth =10335
                                    LayoutCachedHeight =11685
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10125
                                    Top =11835
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3647"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10125
                                    LayoutCachedTop =11835
                                    LayoutCachedWidth =10335
                                    LayoutCachedHeight =12075
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5550
                                    Top =11715
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3649"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5550
                                    LayoutCachedTop =11715
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =11955
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5640
                                    Top =8535
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3650"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =8775
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =4170
                                    Top =8535
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    ForeColor =14270637
                                    Name ="Label3651"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =4170
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =8775
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =9315
                                    Top =7950
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label3652"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =9315
                                    LayoutCachedTop =7950
                                    LayoutCachedWidth =9525
                                    LayoutCachedHeight =8190
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =15045
                                    Top =6915
                                    Width =300
                                    Height =240
                                    FontSize =8
                                    Name ="Label3655"
                                    Caption ="pct."
                                    FontName ="Segoe UI"
                                    Tag ="Acct"
                                    LayoutCachedLeft =15045
                                    LayoutCachedTop =6915
                                    LayoutCachedWidth =15345
                                    LayoutCachedHeight =7155
                                    ForeThemeColorIndex =2
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =15045
                                    Top =7185
                                    Width =300
                                    Height =240
                                    FontSize =8
                                    Name ="Label3656"
                                    Caption ="pct."
                                    FontName ="Segoe UI"
                                    Tag ="Acct"
                                    LayoutCachedLeft =15045
                                    LayoutCachedTop =7185
                                    LayoutCachedWidth =15345
                                    LayoutCachedHeight =7425
                                    ForeThemeColorIndex =1
                                    ForeShade =50.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5640
                                    Top =9540
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3657"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5640
                                    LayoutCachedTop =9540
                                    LayoutCachedWidth =5850
                                    LayoutCachedHeight =9780
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5550
                                    Top =12615
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label3660"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5550
                                    LayoutCachedTop =12615
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =12855
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    FontItalic = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =5670
                                    Top =7320
                                    Width =240
                                    Height =180
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =60
                                    Name ="ProjectedOrActual"
                                    ControlSource ="Assigned_Indicator"
                                    StatusBarText ="P = Assignee Projected, F = Assignee Finalized"
                                    ControlTipText ="P = Assignee Projected, F = Assignee Finalized"

                                    LayoutCachedLeft =5670
                                    LayoutCachedTop =7320
                                    LayoutCachedWidth =5910
                                    LayoutCachedHeight =7500
                                    ForeThemeColorIndex =1
                                    ForeShade =50.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5670
                                    Top =9840
                                    Width =1560
                                    FontSize =9
                                    TabIndex =61
                                    Name ="SyndMiscLabel1"
                                    ControlSource ="SyndMiscLabel1"
                                    StatusBarText ="ATF FSL pmt"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="\"Label1\""
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="This will be the label that will show on reports for this MISC value"

                                    LayoutCachedLeft =5670
                                    LayoutCachedTop =9840
                                    LayoutCachedWidth =7230
                                    LayoutCachedHeight =10080
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5670
                                    Top =10140
                                    Width =1560
                                    FontSize =9
                                    TabIndex =62
                                    Name ="SyndMiscLabel2"
                                    ControlSource ="SyndMiscLabel2"
                                    StatusBarText ="ATF FSL pmt"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="\"Label2\""
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="This will be the label that will show on reports for this MISC value"

                                    LayoutCachedLeft =5670
                                    LayoutCachedTop =10140
                                    LayoutCachedWidth =7230
                                    LayoutCachedHeight =10380
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =9840
                                    Width =1425
                                    FontSize =9
                                    TabIndex =63
                                    Name ="SyndMiscValue1"
                                    ControlSource ="SyndMiscValue1"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="ATF FSL pmt"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="This MISC value wil be included in the SYND report. Negative / Positive will be "
                                        "calulated accordingly"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =9840
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =10080
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =4200
                                    Top =10140
                                    Width =1425
                                    FontSize =9
                                    TabIndex =64
                                    Name ="SyndMiscValue2"
                                    ControlSource ="SyndMiscValue2"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="ATF FSL pmt"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="This MISC value wil be included in the SYND report. Negative / Positive will be "
                                        "calulated accordingly"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =10140
                                    LayoutCachedWidth =5625
                                    LayoutCachedHeight =10380
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1056
                                    Top =9840
                                    Width =1704
                                    Height =240
                                    FontSize =9
                                    BackColor =15527148
                                    Name ="Label4084"
                                    Caption ="Misc Value / Label 1 "
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1056
                                    LayoutCachedTop =9840
                                    LayoutCachedWidth =2760
                                    LayoutCachedHeight =10080
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =5
                                    ColumnEnd =5
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1056
                                    Top =10160
                                    Width =1704
                                    Height =240
                                    FontSize =9
                                    BackColor =15527148
                                    Name ="Label4082"
                                    Caption ="Misc Value / Label 2 "
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =1056
                                    LayoutCachedTop =10160
                                    LayoutCachedWidth =2760
                                    LayoutCachedHeight =10400
                                    RowStart =15
                                    RowEnd =15
                                    ColumnStart =5
                                    ColumnEnd =5
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =2940
                                    Top =9840
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    Name ="Label4089"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =2940
                                    LayoutCachedTop =9840
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =10080
                                    ForeThemeColorIndex =1
                                    ForeShade =75.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =2940
                                    Top =10140
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    Name ="Label4090"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =2940
                                    LayoutCachedTop =10140
                                    LayoutCachedWidth =4140
                                    LayoutCachedHeight =10380
                                    ForeThemeColorIndex =1
                                    ForeShade =75.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6000
                                    Top =12285
                                    Width =855
                                    Height =555
                                    FontSize =9
                                    TabIndex =65
                                    BorderColor =0
                                    ForeColor =1643706
                                    Name ="SyndMsg"
                                    FontName ="Segoe UI"
                                    TextFormat =1

                                    LayoutCachedLeft =6000
                                    LayoutCachedTop =12285
                                    LayoutCachedWidth =6855
                                    LayoutCachedHeight =12840
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =5700
                                    Top =8835
                                    Width =930
                                    Height =285
                                    FontSize =9
                                    TabIndex =66
                                    Name ="PerDiemCalculated"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"
                                    ControlTipText ="Calculated Per Diem based on unit level calculations"

                                    LayoutCachedLeft =5700
                                    LayoutCachedTop =8835
                                    LayoutCachedWidth =6630
                                    LayoutCachedHeight =9120
                                    RowStart =9
                                    RowEnd =9
                                    ColumnStart =3
                                    ColumnEnd =3
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =12150
                                    Top =6915
                                    Width =2205
                                    Height =240
                                    FontSize =8
                                    Name ="Label3617"
                                    Caption ="Interim Interest Rate @ "
                                    FontName ="Segoe UI"
                                    Tag ="Acct"
                                    LayoutCachedLeft =12150
                                    LayoutCachedTop =6915
                                    LayoutCachedWidth =14355
                                    LayoutCachedHeight =7155
                                    ForeThemeColorIndex =2
                                End
                                Begin CommandButton
                                    DisplayWhen =2
                                    OverlapFlags =247
                                    Left =15495
                                    Top =6600
                                    Width =1590
                                    Height =278
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =67
                                    Name ="SyndCheckExcel"
                                    Caption =" Sch Synd Check"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                                        0xddddd87626dddddd
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    LeftPadding =60
                                    TopPadding =15
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =15495
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =17085
                                    LayoutCachedHeight =6878
                                    PictureCaptionArrangement =5
                                    ForeThemeColorIndex =6
                                    ForeShade =50.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =-2147483633
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    HoverThemeColorIndex =6
                                    HoverShade =75.0
                                    QuickStyle =30
                                    QuickStyleMask =-629
                                    Overlaps =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =11460
                                    Top =8670
                                    Width =7920
                                    Height =3558
                                    TabIndex =68
                                    Name ="SynMemo"
                                    ControlSource ="SynMemo"
                                    Tag ="StayOpen"

                                    LayoutCachedLeft =11460
                                    LayoutCachedTop =8670
                                    LayoutCachedWidth =19380
                                    LayoutCachedHeight =12228
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =11445
                                    Top =8505
                                    Width =7740
                                    Height =180
                                    Name ="Image3658"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =11445
                                    LayoutCachedTop =8505
                                    LayoutCachedWidth =19185
                                    LayoutCachedHeight =8685
                                    TabIndex =69
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =11445
                                    Top =8205
                                    Width =7920
                                    Height =360
                                    FontWeight =700
                                    Name ="Label3659"
                                    Caption ="Syndication Memo"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =11445
                                    LayoutCachedTop =8205
                                    LayoutCachedWidth =19365
                                    LayoutCachedHeight =8565
                                    ColumnStart =9
                                    ColumnEnd =10
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8820
                                    Top =12105
                                    Width =1260
                                    Height =285
                                    FontSize =9
                                    TabIndex =70
                                    Name ="InterimExpense"
                                    ControlSource ="=DLookUp(\"GroupInterimExp\",\"InterimFunding_Days_Rate_Group\",\"GroupRef=\" & "
                                        "[SGrpID])"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =8820
                                    LayoutCachedTop =12105
                                    LayoutCachedWidth =10080
                                    LayoutCachedHeight =12390
                                    RowStart =13
                                    RowEnd =13
                                    ColumnStart =3
                                    ColumnEnd =3
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =7305
                                    Top =12105
                                    Width =1425
                                    Height =285
                                    FontSize =9
                                    Name ="Label4469"
                                    Caption ="Interim Expense"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =7305
                                    LayoutCachedTop =12105
                                    LayoutCachedWidth =8730
                                    LayoutCachedHeight =12390
                                    RowStart =13
                                    RowEnd =13
                                    ColumnEnd =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10125
                                    Top =12150
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    Name ="Label4470"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10125
                                    LayoutCachedTop =12150
                                    LayoutCachedWidth =10335
                                    LayoutCachedHeight =12390
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin Label
                                    Visible = NotDefault
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =2
                                    Left =17400
                                    Top =6180
                                    Width =1320
                                    Height =285
                                    FontSize =10
                                    Name ="Label4509"
                                    Caption ="RV Guar Type"
                                    LayoutCachedLeft =17400
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =18720
                                    LayoutCachedHeight =6465
                                    BackThemeColorIndex =9
                                    BackShade =75.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =17400
                                    Top =6390
                                    Width =1320
                                    Height =315
                                    TabIndex =71
                                    Name ="RVGuarType"
                                    ControlSource ="RVGuarType"
                                    RowSourceType ="Value List"
                                    RowSource ="0;None;1;FA;2;OEM;3;Both"
                                    ColumnWidths ="0;1440"
                                    StatusBarText ="RV Guar Type"
                                    Format ="General Number"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =17400
                                    LayoutCachedTop =6390
                                    LayoutCachedWidth =18720
                                    LayoutCachedHeight =6705
                                    ColumnStart =1
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =2
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =8565
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =72
                                    Name ="SpreadIRR"
                                    ControlSource ="=[YieldIRR]-[IndexRate]"
                                    Format ="Percent"
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =8565
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =8850
                                    RowStart =11
                                    RowEnd =11
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =6960
                                            Top =8550
                                            Width =1200
                                            Height =285
                                            FontSize =9
                                            Name ="Label4518"
                                            Caption ="IRR-PT Spread"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =6960
                                            LayoutCachedTop =8550
                                            LayoutCachedWidth =8160
                                            LayoutCachedHeight =8835
                                            RowStart =11
                                            RowEnd =11
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10665
                                    Top =8595
                                    Width =210
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label4513"
                                    Caption ="Δ"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10665
                                    LayoutCachedTop =8595
                                    LayoutCachedWidth =10875
                                    LayoutCachedHeight =8835
                                    ForeThemeColorIndex =4
                                    ForeShade =75.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =8235
                                    Width =1140
                                    Height =285
                                    FontSize =9
                                    TabIndex =73
                                    Name ="YieldIRR"
                                    ControlSource ="YieldIRR"
                                    Format ="Percent"
                                    FontName ="Segoe UI"
                                    Tag ="TrackUpdate"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =8235
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =8520
                                    RowStart =12
                                    RowEnd =12
                                    ColumnStart =7
                                    ColumnEnd =7
                                    BackThemeColorIndex =8
                                    BackTint =40.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7260
                                            Top =8220
                                            Width =900
                                            Height =285
                                            FontSize =9
                                            Name ="Label4519"
                                            Caption ="IRR Yield"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =7260
                                            LayoutCachedTop =8220
                                            LayoutCachedWidth =8160
                                            LayoutCachedHeight =8505
                                            RowStart =12
                                            RowEnd =12
                                            ColumnStart =5
                                            ColumnEnd =5
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =8220
                                    Top =8220
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    ForeColor =6405021
                                    Name ="Label4520"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =8220
                                    LayoutCachedWidth =9420
                                    LayoutCachedHeight =8460
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =8220
                                    Top =8580
                                    Width =1200
                                    Height =240
                                    FontSize =8
                                    ForeColor =6405021
                                    Name ="Label4521"
                                    Caption ="────────►"
                                    FontName ="Arial"
                                    LayoutCachedLeft =8220
                                    LayoutCachedTop =8580
                                    LayoutCachedWidth =9420
                                    LayoutCachedHeight =8820
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =9450
                                    Top =6285
                                    Width =1140
                                    Height =255
                                    FontSize =9
                                    TabIndex =74
                                    Name ="SGrpIndex"
                                    ControlSource ="SGrpIndex"
                                    RowSourceType ="Value List"
                                    RowSource ="\"LIBOR\";\"SOFR\";\"T-BILL\""
                                    FontName ="Segoe UI"

                                    LayoutCachedLeft =9450
                                    LayoutCachedTop =6285
                                    LayoutCachedWidth =10590
                                    LayoutCachedHeight =6540
                                    RowStart =7
                                    RowEnd =7
                                    ColumnStart =6
                                    ColumnEnd =6
                                    ForeThemeColorIndex =0
                                    ForeTint =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="SyndicationNotes"
                            Caption ="Documentation"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureType =2
                                    Left =630
                                    Top =6000
                                    Width =4770
                                    Height =240
                                    Name ="Image3575"
                                    Picture ="3_WhiteShadow"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =630
                                    LayoutCachedTop =6000
                                    LayoutCachedWidth =5400
                                    LayoutCachedHeight =6240
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =630
                                    Top =6180
                                    Width =1860
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3233"
                                    Caption ="      Residual Guarantee"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =2490
                                    LayoutCachedHeight =6435
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =885
                                    Top =6510
                                    TabIndex =1
                                    Name ="RGDCasualtyYN"
                                    ControlSource ="RGDCasualtyYN"
                                    Tag ="RGD"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =6510
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =6750
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =1070
                                            Top =6480
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3542"
                                            Caption ="Casualty"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =6480
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =6750
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =885
                                    Top =6765
                                    TabIndex =2
                                    Name ="RGDExchangeITYN"
                                    ControlSource ="RGDExchangeITYN"
                                    Tag ="RGD"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =6765
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =7005
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =1070
                                            Top =6735
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3545"
                                            Caption ="ExchangeIT"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =6735
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =7005
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =6240
                                    Width =240
                                    Height =180
                                    TabIndex =3
                                    Name ="ResidualGuaranteeDocYN"
                                    ControlSource ="ResidualGuaranteeDocYN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =6420
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =2490
                                    Top =6180
                                    Width =1935
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3571"
                                    Caption ="      Shortfall Agreement"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =2490
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =4425
                                    LayoutCachedHeight =6435
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2550
                                    Top =6240
                                    Width =240
                                    Height =180
                                    TabIndex =4
                                    Name ="ShortFallAgreementYN"
                                    ControlSource ="ShortFallAgreementYN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =2550
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =2790
                                    LayoutCachedHeight =6420
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =630
                                    Top =7740
                                    Width =2145
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3573"
                                    Caption ="      Indemnity Letter Agmt"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =7740
                                    LayoutCachedWidth =2775
                                    LayoutCachedHeight =7995
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =7800
                                    Width =240
                                    Height =180
                                    TabIndex =5
                                    Name ="IndemnificationDocYN"
                                    ControlSource ="IndemnificationDocYN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =7800
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =7980
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =885
                                    Top =7050
                                    TabIndex =6
                                    Name ="RGDRepurchaseTN"
                                    ControlSource ="RGDRepurchaseTN"
                                    Tag ="RGD"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =7050
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =7290
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =1070
                                            Top =7020
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3547"
                                            Caption ="Repurchase"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =7020
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =7290
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =885
                                    Top =7305
                                    TabIndex =7
                                    Name ="RGDTerminationYN"
                                    ControlSource ="RGDTerminationYN"
                                    Tag ="RGD"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =7305
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =7545
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =1070
                                            Top =7275
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3549"
                                            Caption ="Termination"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =7275
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =7545
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =2760
                                    Top =6510
                                    TabIndex =8
                                    Name ="SFACasualtyYN"
                                    ControlSource ="SFACasualtyYN"
                                    Tag ="SFA"

                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =6510
                                    LayoutCachedWidth =3020
                                    LayoutCachedHeight =6750
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =2945
                                            Top =6480
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3555"
                                            Caption ="Casualty"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =2945
                                            LayoutCachedTop =6480
                                            LayoutCachedWidth =4040
                                            LayoutCachedHeight =6750
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =2760
                                    Top =6765
                                    TabIndex =9
                                    Name ="SFAExchangeITYN"
                                    ControlSource ="SFAExchangeITYN"
                                    Tag ="SFA"

                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =6765
                                    LayoutCachedWidth =3020
                                    LayoutCachedHeight =7005
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =2945
                                            Top =6735
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3557"
                                            Caption ="ExchangeIT"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =2945
                                            LayoutCachedTop =6735
                                            LayoutCachedWidth =4040
                                            LayoutCachedHeight =7005
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =2760
                                    Top =7050
                                    TabIndex =10
                                    Name ="SFARepurchaseTN"
                                    ControlSource ="SFARepurchaseTN"
                                    Tag ="SFA"

                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =7050
                                    LayoutCachedWidth =3020
                                    LayoutCachedHeight =7290
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =2945
                                            Top =7020
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3559"
                                            Caption ="Repurchase"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =2945
                                            LayoutCachedTop =7020
                                            LayoutCachedWidth =4040
                                            LayoutCachedHeight =7290
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =2760
                                    Top =7305
                                    TabIndex =11
                                    Name ="SFATerminationYN"
                                    ControlSource ="SFATerminationYN"
                                    Tag ="SFA"

                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =7305
                                    LayoutCachedWidth =3020
                                    LayoutCachedHeight =7545
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =2945
                                            Top =7275
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3561"
                                            Caption ="Termination"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =2945
                                            LayoutCachedTop =7275
                                            LayoutCachedWidth =4040
                                            LayoutCachedHeight =7545
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =885
                                    Top =8070
                                    TabIndex =12
                                    Name ="IDFLateFeesYN"
                                    ControlSource ="IDFLateFeesYN"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =8070
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =8310
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =1070
                                            Top =8040
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3563"
                                            Caption ="Late Fees"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =8040
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =8310
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =885
                                    Top =8325
                                    TabIndex =13
                                    Name ="IDFMileageYN"
                                    ControlSource ="IDFMileageYN"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =8325
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =8565
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =1070
                                            Top =8295
                                            Width =1095
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3565"
                                            Caption ="Mileage"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =8295
                                            LayoutCachedWidth =2165
                                            LayoutCachedHeight =8565
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =630
                                    Top =5700
                                    Width =4770
                                    Height =360
                                    FontWeight =700
                                    Name ="Label2872"
                                    Caption ="Special Assignment Agreements"
                                    FontName ="Segoe UI"
                                    GridlineStyleBottom =1
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =5700
                                    LayoutCachedWidth =5400
                                    LayoutCachedHeight =6060
                                    ColumnStart =9
                                    ColumnEnd =10
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =630
                                    Top =10860
                                    Width =2880
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3576"
                                    Caption ="      Omnibus Agreement "
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =10860
                                    LayoutCachedWidth =3510
                                    LayoutCachedHeight =11115
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =10920
                                    Width =240
                                    Height =180
                                    TabIndex =14
                                    Name ="OmnibusAgmtYN"
                                    ControlSource ="OmnibusAgmtYN"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =10920
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =11100
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =630
                                    Top =10500
                                    Width =2880
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3578"
                                    Caption ="      \011Loan && Security Agreement "
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =10500
                                    LayoutCachedWidth =3510
                                    LayoutCachedHeight =10755
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =10560
                                    Width =240
                                    Height =180
                                    TabIndex =15
                                    Name ="LoanSecurityAgmtYN"
                                    ControlSource ="LoanSecurityAgmtYN"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =10560
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =10740
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =630
                                    Top =10140
                                    Width =2880
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3580"
                                    Caption ="      \0113rd Party Pay Agent Assignment"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =10140
                                    LayoutCachedWidth =3510
                                    LayoutCachedHeight =10395
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =10200
                                    Width =240
                                    Height =180
                                    TabIndex =16
                                    Name ="ThirdPartyPayAssgnmtYN"
                                    ControlSource ="ThirdPartyPayAssgnmtYN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =10200
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =10380
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =630
                                    Top =11580
                                    Width =3899
                                    Height =300
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3582"
                                    Caption ="      \011Custom Document Name"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =11580
                                    LayoutCachedWidth =4529
                                    LayoutCachedHeight =11880
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =690
                                    Top =11640
                                    Width =240
                                    Height =285
                                    TabIndex =17
                                    Name ="CustomDocumentNameYN"
                                    ControlSource ="CustomDocumentNameYN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =11640
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =11925
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =720
                                    Top =11940
                                    Width =3720
                                    Height =720
                                    FontSize =9
                                    TabIndex =18
                                    BorderColor =2838138
                                    Name ="CustomDocumentName"
                                    ControlSource ="CustomDocumentName"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"

                                    LayoutCachedLeft =720
                                    LayoutCachedTop =11940
                                    LayoutCachedWidth =4440
                                    LayoutCachedHeight =12660
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =630
                                            Top =11880
                                            Width =3899
                                            Height =840
                                            FontSize =9
                                            BorderColor =2838138
                                            Name ="CustomDocumentNameLBL"
                                            Caption ="  "
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =630
                                            LayoutCachedTop =11880
                                            LayoutCachedWidth =4529
                                            LayoutCachedHeight =12720
                                            RowStart =1
                                            RowEnd =1
                                            ColumnStart =9
                                            ColumnEnd =9
                                            BackThemeColorIndex =1
                                            BackShade =95.0
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =630
                                    Top =11220
                                    Width =2880
                                    Height =255
                                    FontSize =9
                                    BorderColor =2838138
                                    Name ="Label3586"
                                    Caption ="      Progress Payment Agreement"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =630
                                    LayoutCachedTop =11220
                                    LayoutCachedWidth =3510
                                    LayoutCachedHeight =11475
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =9
                                    ColumnEnd =9
                                    BackThemeColorIndex =1
                                    BackShade =95.0
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =690
                                    Top =11280
                                    Width =240
                                    Height =180
                                    TabIndex =19
                                    Name ="ProgressPymtAgmtYN"
                                    ControlSource ="ProgressPymtAgmtYN"

                                    LayoutCachedLeft =690
                                    LayoutCachedTop =11280
                                    LayoutCachedWidth =930
                                    LayoutCachedHeight =11460
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =10
                                    ColumnEnd =10
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =885
                                    Top =8610
                                    TabIndex =20
                                    Name ="LesseeDefaultsAll"
                                    ControlSource ="LesseeDefaultsAll"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =8610
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =8850
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =1070
                                            Top =8580
                                            Width =1620
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3665"
                                            Caption ="\011Lessee Defaults ALL"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =8580
                                            LayoutCachedWidth =2690
                                            LayoutCachedHeight =8850
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =885
                                    Top =8865
                                    TabIndex =21
                                    Name ="LesseeDefaultsCertain"
                                    ControlSource ="LesseeDefaultsCertain"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =8865
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =9105
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =1070
                                            Top =8835
                                            Width =1890
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="LesseeDefaultsCertainLbl"
                                            Caption ="Lessee Defaults Certain"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =8835
                                            LayoutCachedWidth =2960
                                            LayoutCachedHeight =9105
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =255
                                    Left =885
                                    Top =9150
                                    TabIndex =22
                                    Name ="LesseeObligationsAll"
                                    ControlSource ="LesseeObligationsAll"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =9150
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =9390
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =1070
                                            Top =9120
                                            Width =1890
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3669"
                                            Caption ="Lessee Obligations ALL"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =9120
                                            LayoutCachedWidth =2960
                                            LayoutCachedHeight =9390
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =885
                                    Top =9405
                                    TabIndex =23
                                    Name ="LesseeObligationsCertain"
                                    ControlSource ="LesseeObligationsCertain"
                                    Tag ="IDF"

                                    LayoutCachedLeft =885
                                    LayoutCachedTop =9405
                                    LayoutCachedWidth =1145
                                    LayoutCachedHeight =9645
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =1070
                                            Top =9375
                                            Width =2160
                                            Height =270
                                            FontSize =9
                                            BorderColor =2838138
                                            ForeColor =8210719
                                            Name ="Label3671"
                                            Caption ="Lessee Obligations Certain"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =1070
                                            LayoutCachedTop =9375
                                            LayoutCachedWidth =3230
                                            LayoutCachedHeight =9645
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =5760
                                    Top =5400
                                    Width =4920
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =24
                                    Name ="AddDocLink"
                                    Caption ="Add Another Document Link"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =5760
                                    LayoutCachedTop =5400
                                    LayoutCachedWidth =10680
                                    LayoutCachedHeight =5700
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackThemeColorIndex =6
                                    BackShade =75.0
                                    BorderThemeColorIndex =7
                                    HoverThemeColorIndex =6
                                    HoverShade =50.0
                                    PressedThemeColorIndex =6
                                    PressedShade =75.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =33
                                    QuickStyleMask =-565
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5700
                                    Top =5760
                                    Width =13680
                                    Height =7320
                                    TabIndex =25
                                    Name ="DocLinksSub"
                                    SourceObject ="Form.DocLinksSub"

                                    LayoutCachedLeft =5700
                                    LayoutCachedTop =5760
                                    LayoutCachedWidth =19380
                                    LayoutCachedHeight =13080
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="BillingRules"
                            Caption ="Billing Rules"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Rectangle
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    Left =305
                                    Top =5570
                                    Width =18090
                                    Height =7335
                                    Name ="Box447"
                                    LayoutCachedLeft =305
                                    LayoutCachedTop =5570
                                    LayoutCachedWidth =18395
                                    LayoutCachedHeight =12905
                                    BorderThemeColorIndex =6
                                    BorderShade =75.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =450
                                    Top =5400
                                    Width =1590
                                    Height =360
                                    FontSize =14
                                    FontWeight =700
                                    Name ="Label448"
                                    Caption =" Billing Rules"
                                    LayoutCachedLeft =450
                                    LayoutCachedTop =5400
                                    LayoutCachedWidth =2040
                                    LayoutCachedHeight =5760
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    ForeThemeColorIndex =2
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =450
                                    Top =5820
                                    Width =1260
                                    Height =255
                                    FontSize =9
                                    FontWeight =700
                                    Name ="Label63"
                                    Caption ="Bill To Address"
                                    FontName ="Segoe UI"
                                    RightPadding =0
                                    LayoutCachedLeft =450
                                    LayoutCachedTop =5820
                                    LayoutCachedWidth =1710
                                    LayoutCachedHeight =6075
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =8
                                    ListWidth =8640
                                    Left =1740
                                    Top =5820
                                    Width =4170
                                    Height =255
                                    FontSize =9
                                    Name ="BillToAddressID"
                                    ControlSource ="BillToAddressID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, [LocationStreet] & Chr(13) & Chr(10) & IIf(IsNull(["
                                        "LAddrLine1]),Null,[Locations].[LAddrLine1] & Chr(13) & Chr(10)) & IIf(IsNull([LA"
                                        "ddrLine2]),Null,[Locations].[LAddrLine2] & Chr(13) & Chr(10)) & [LCity] & \" \" "
                                        "& [LState] & \" \" & [LZip] AS BillingAddress, Locations.LocationType, Locations"
                                        ".LocationName, Locations.LAddrLine2, Locations.Locationsn, Locations.LCity, Loca"
                                        "tions.LState FROM Locations WHERE (((Locations.LocationType)='Billing' Or (Locat"
                                        "ions.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)=[Forms]![GroupForm"
                                        "]![ClientGroupID]) AND ((Locations.LocObsolete)<>Yes)) ORDER BY Locations.Locati"
                                        "onType, Locations.LocationName;"
                                    ColumnWidths ="0;0;720;2160;2160;1296;1296;360"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="This is the 'Bill To' address that will appear on ALL Invoices for this group"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =1740
                                    LayoutCachedTop =5820
                                    LayoutCachedWidth =5910
                                    LayoutCachedHeight =6075
                                    ColumnStart =1
                                    ColumnEnd =7
                                    BorderThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    SpecialEffect =3
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =6396
                                    Top =5820
                                    Width =3432
                                    Height =360
                                    FontWeight =700
                                    Name ="Label908"
                                    Caption ="FA Bill for Assignee & Invoicing"
                                    FontName ="Segoe UI"
                                    GroupTable =25
                                    LayoutCachedLeft =6396
                                    LayoutCachedTop =5820
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =6180
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListRows =3
                                    ListWidth =1080
                                    Left =8712
                                    Top =6252
                                    Width =1116
                                    FontSize =9
                                    TabIndex =1
                                    Name ="SgrpCustPmt"
                                    ControlSource ="SgrpCustPmt"
                                    RowSourceType ="Value List"
                                    RowSource ="1;Check;2;ACH;3;Fed Wire"
                                    ColumnWidths ="0;1080"
                                    StatusBarText ="Customer Pays bank by check, Fed wire or ACH"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Select a Payment Method the Client Will Use"
                                    GroupTable =25
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =6252
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =6492
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =6252
                                            Width =2280
                                            Height =240
                                            FontSize =9
                                            Name ="Label409"
                                            Caption =" Pmt Method"
                                            FontName ="Segoe UI"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =6252
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =6492
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =3
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    Left =8712
                                    Top =6552
                                    Width =1116
                                    Height =252
                                    TabIndex =2
                                    Name ="SgrpFA4Assign"
                                    ControlSource ="SgrpFA4Assign"
                                    StatusBarText ="FA bills customer for bank after assignment"
                                    GroupTable =25

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =6552
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =6804
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =6552
                                            Width =2280
                                            Height =252
                                            FontSize =9
                                            Name ="Label408"
                                            Caption ="FA Bills (check for yes)"
                                            FontName ="Segoe UI"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =6552
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =6804
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =3
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =495
                                    Top =6210
                                    Width =2895
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label657"
                                    Caption ="┌─Bill To Address Preview ─────┐"
                                    FontName ="Segoe UI"
                                    RightPadding =0
                                    LayoutCachedLeft =495
                                    LayoutCachedTop =6210
                                    LayoutCachedWidth =3390
                                    LayoutCachedHeight =6450
                                    RowStart =1
                                    RowEnd =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =2
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =525
                                    Top =6405
                                    Width =5370
                                    Height =1620
                                    FontSize =9
                                    TabIndex =3
                                    BorderColor =0
                                    ForeColor =-2147483617
                                    Name ="BillingAddressPreview"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    ControlTipText ="This is the 'Bill To' address that will appear on ALL Bills and  Invoices"

                                    LayoutCachedLeft =525
                                    LayoutCachedTop =6405
                                    LayoutCachedWidth =5895
                                    LayoutCachedHeight =8025
                                    BackThemeColorIndex =1
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8712
                                    Top =7920
                                    Width =1116
                                    Height =276
                                    FontSize =9
                                    TabIndex =4
                                    Name ="FA3rdPartyPayingAgent"
                                    ControlSource ="FA3rdPartyPayingAgent"
                                    RowSourceType ="Value List"
                                    RowSource ="No;Needed;Pending;Done"
                                    FontName ="Segoe UI"
                                    OnClick ="[Event Procedure]"
                                    GroupTable =25

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =7920
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =8196
                                    RowStart =6
                                    RowEnd =6
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =7920
                                            Width =2280
                                            Height =276
                                            FontSize =10
                                            Name ="Label2778"
                                            Caption ="FA 3rd Party Paying Agent:"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =7920
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =8196
                                            RowStart =6
                                            RowEnd =6
                                            LayoutGroup =3
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8712
                                    Top =7572
                                    Width =1116
                                    Height =276
                                    FontSize =9
                                    TabIndex =5
                                    Name ="FACollecting"
                                    ControlSource ="FACollecting"
                                    RowSourceType ="Value List"
                                    RowSource ="No;Yes"
                                    FontName ="Segoe UI"
                                    GroupTable =25

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =7572
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =7848
                                    RowStart =5
                                    RowEnd =5
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =7572
                                            Width =2280
                                            Height =276
                                            FontSize =10
                                            Name ="Label2773"
                                            Caption ="FA Collecting"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =7572
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =7848
                                            RowStart =5
                                            RowEnd =5
                                            LayoutGroup =3
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8712
                                    Top =7224
                                    Width =1116
                                    Height =276
                                    FontSize =9
                                    TabIndex =6
                                    Name ="FAInvoicing"
                                    ControlSource ="FAInvoicing"
                                    RowSourceType ="Value List"
                                    RowSource ="No;Yes"
                                    FontName ="Segoe UI"
                                    GroupTable =25

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =7224
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =7500
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =7224
                                            Width =2280
                                            Height =276
                                            FontSize =10
                                            Name ="Label2768"
                                            Caption ="FA Invoicing"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =7224
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =7500
                                            RowStart =4
                                            RowEnd =4
                                            LayoutGroup =3
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =8712
                                    Top =6876
                                    Width =1116
                                    Height =276
                                    FontSize =9
                                    TabIndex =7
                                    Name ="FAServicer"
                                    ControlSource ="FAServicer"
                                    RowSourceType ="Value List"
                                    RowSource ="No;Yes"
                                    FontName ="Segoe UI"
                                    GroupTable =25

                                    LayoutCachedLeft =8712
                                    LayoutCachedTop =6876
                                    LayoutCachedWidth =9828
                                    LayoutCachedHeight =7152
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =25
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =6396
                                            Top =6876
                                            Width =2280
                                            Height =276
                                            FontSize =10
                                            Name ="Label2763"
                                            Caption ="FA Servicer"
                                            GroupTable =25
                                            RightPadding =0
                                            LayoutCachedLeft =6396
                                            LayoutCachedTop =6876
                                            LayoutCachedWidth =8676
                                            LayoutCachedHeight =7152
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =3
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =25
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =540
                                    Top =8460
                                    Width =5355
                                    Height =4275
                                    FontSize =9
                                    TabIndex =8
                                    ForeColor =-2147483617
                                    Name ="SpecialRemitAddrGrp"
                                    ControlSource ="SpecialRemitAddrGrp"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    ControlTipText ="This is the 'Remit To' address that will appear on ALL Bills and  Invoices"
                                    TextFormat =1

                                    LayoutCachedLeft =540
                                    LayoutCachedTop =8460
                                    LayoutCachedWidth =5895
                                    LayoutCachedHeight =12735
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =95.0
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    Left =540
                                    Top =8160
                                    Width =1110
                                    Height =255
                                    TabIndex =9
                                    Name ="UseSpecialRemitGrp"
                                    ControlSource ="UseSpecialRemitGrp"
                                    StatusBarText ="FA bills customer for bank after assignment"
                                    OnClick ="[Event Procedure]"
                                    GroupTable =50

                                    LayoutCachedLeft =540
                                    LayoutCachedTop =8160
                                    LayoutCachedWidth =1650
                                    LayoutCachedHeight =8415
                                    LayoutGroup =7
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =50
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =720
                                    Top =8160
                                    Width =2400
                                    Height =255
                                    FontSize =9
                                    Name ="Label4065"
                                    Caption ="Use Special Remit To Address"
                                    FontName ="Segoe UI"
                                    GroupTable =51
                                    RightPadding =0
                                    LayoutCachedLeft =720
                                    LayoutCachedTop =8160
                                    LayoutCachedWidth =3120
                                    LayoutCachedHeight =8415
                                    LayoutGroup =8
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =51
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10185
                                    Top =6405
                                    Width =7950
                                    Height =1725
                                    FontSize =9
                                    TabIndex =10
                                    BorderColor =0
                                    ForeColor =-2147483617
                                    Name ="BillingInstructions"
                                    ControlSource ="BillingInstructions"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    ControlTipText ="These are Billing Instructions that apply to this group alone"

                                    LayoutCachedLeft =10185
                                    LayoutCachedTop =6405
                                    LayoutCachedWidth =18135
                                    LayoutCachedHeight =8130
                                    BackThemeColorIndex =1
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =10185
                                            Top =6210
                                            Width =2895
                                            Height =240
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label4393"
                                            Caption ="┌─ Billing Instructions ─────┐"
                                            FontName ="Segoe UI"
                                            RightPadding =0
                                            LayoutCachedLeft =10185
                                            LayoutCachedTop =6210
                                            LayoutCachedWidth =13080
                                            LayoutCachedHeight =6450
                                            RowStart =1
                                            RowEnd =1
                                            BackThemeColorIndex =1
                                            BorderThemeColorIndex =2
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10185
                                    Top =8535
                                    Width =7950
                                    Height =1725
                                    FontSize =9
                                    TabIndex =11
                                    BorderColor =0
                                    ForeColor =-2147483617
                                    Name ="Text4396"
                                    ControlSource ="TaxInstructions"
                                    FontName ="Segoe UI"
                                    Tag ="StayOpen"
                                    ControlTipText ="These are Billing Instructions that apply to this group alone"

                                    LayoutCachedLeft =10185
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =18135
                                    LayoutCachedHeight =10260
                                    BackThemeColorIndex =1
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =10185
                                            Top =8340
                                            Width =2895
                                            Height =240
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label4397"
                                            Caption ="┌─ Tax Instructions ─────┐"
                                            FontName ="Segoe UI"
                                            RightPadding =0
                                            LayoutCachedLeft =10185
                                            LayoutCachedTop =8340
                                            LayoutCachedWidth =13080
                                            LayoutCachedHeight =8580
                                            RowStart =1
                                            RowEnd =1
                                            BackThemeColorIndex =1
                                            BorderThemeColorIndex =2
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =10200
                                    Top =10740
                                    Width =3780
                                    FontWeight =700
                                    TabIndex =12
                                    Name ="MakePaiLetter"
                                    Caption ="  Make Wire Payment Instructions"
                                    OnClick ="[Event Procedure]"
                                    Tag ="UAT"
                                    LeftPadding =150
                                    TopPadding =105
                                    RightPadding =165
                                    BottomPadding =165
                                    ImageData = Begin
                                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000913f10ef913f10ff913f10ff ,
                                        0x913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff ,
                                        0x913f10ff913f10ef000000000000000000000000913f10ff913f10ff913f10ff ,
                                        0x913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff913f10ff ,
                                        0x913f10ff913f10ff000000000000000000000000481f08ff481f08ff481f08ff ,
                                        0x481f08ff481f08ff481f08ff481f08ff6d2f0cff913f10ff913f10ff913f10ff ,
                                        0x913f10ff913f10ffbd5a18efbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffb65617ff481f08ff913f10ff913f10ff913f10ff ,
                                        0x913f10ff913f10ffbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffbd5a18ff5e2d0cffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffbd5a18ffe2b79affffffffffce8352ffce8352ff ,
                                        0xffffffffe2b79affbd5a18ffbd5a18ff5e2d0cffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffbd5a18ffefd6c5ffffffffffdead8cffdaa27dff ,
                                        0xffffffffefd6c5ffbd5a18ffbd5a18ff5e2d0cffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffbd5a18fffffffffff3e0d4ffeacbb7ffeacbb7ff ,
                                        0xf3e0d4ffffffffffbd5a18ffbd5a18ff5e2d0cffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffce8352ffffffffffd6986ffff7eae2fffbf5f1ff ,
                                        0xd6986fffffffffffce8352ffbd5a18ff693e15ffd37c2bffd37c2bffd37c2bff ,
                                        0xd37c2bffd37c2bffbd5a18ffdaa27dfffbf5f1ffc16426ffffffffffffffffff ,
                                        0xc56f35fff7eae2ffdaa27dffbd5a18ff693e15ffd37c2bffd37c2bffd37c2bff ,
                                        0xd37c2bffd37c2bffbd5a18ffeacbb7ffefd6c5ffbd5a18fff3e0d4fff7eae2ff ,
                                        0xbd5a18ffefd6c5ffeacbb7ffbd5a18ff693e15ffd37c2bffd37c2bffd37c2bff ,
                                        0xd37c2bffd37c2bffbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffbd5a18ff704216ffd37c2bffd37c2bffd37c2bff ,
                                        0xd37c2bffd37c2bffbd5a18efbd5a18ffbd5a18ffbd5a18ffbd5a18ffbd5a18ff ,
                                        0xbd5a18ffbd5a18ffbd5a18ffc05f1bffeea541ffeea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ff000000000000000000000000eea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ff0000000000000000eea54110eea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ff0000000000000000eea54110eea541efeea541ffeea541ff ,
                                        0xeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ffeea541ff ,
                                        0xeea541ffeea541ef
                                    End

                                    LayoutCachedLeft =10200
                                    LayoutCachedTop =10740
                                    LayoutCachedWidth =13980
                                    LayoutCachedHeight =11100
                                    PictureCaptionArrangement =5
                                    Alignment =1
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =10319446
                                    BorderThemeColorIndex =1
                                    HoverColor =10319446
                                    PressedColor =10319446
                                    HoverForeColor =6108695
                                    PressedForeThemeColorIndex =1
                                    Shadow =5
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =45
                            Top =5160
                            Width =19395
                            Height =9030
                            Name ="ExitOptions"
                            OnMouseMove ="[Event Procedure]"
                            Caption ="EXIT Options"
                            LayoutCachedLeft =45
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =45
                                    Top =5625
                                    Width =18944
                                    Height =7446
                                    Name ="ExitOptionsGroup_Sub"
                                    SourceObject ="Form.ExitOptionsGroup_Sub"
                                    LinkChildFields ="GroupID"
                                    LinkMasterFields ="SGrpID"

                                    LayoutCachedLeft =45
                                    LayoutCachedTop =5625
                                    LayoutCachedWidth =18989
                                    LayoutCachedHeight =13071
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =151
                                            Top =5325
                                            Width =4065
                                            Height =360
                                            FontSize =14
                                            ForeColor =12349952
                                            Name ="Label3473"
                                            Caption ="EXIT Options"
                                            LayoutCachedLeft =151
                                            LayoutCachedTop =5325
                                            LayoutCachedWidth =4216
                                            LayoutCachedHeight =5685
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =12930
                                    Top =5340
                                    Width =1875
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="OpenExitOptions"
                                    Caption =" Add /Edit Group Options"
                                    OnClick ="[Event Procedure]"
                                    Tag ="UAT"
                                    LeftPadding =90
                                    RightPadding =105
                                    BottomPadding =165
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =12930
                                    LayoutCachedTop =5340
                                    LayoutCachedWidth =14805
                                    LayoutCachedHeight =5580
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =8210719
                                    BorderThemeColorIndex =4
                                    BorderShade =75.0
                                    HoverColor =14270637
                                    PressedColor =8210719
                                    HoverForeColor =6108695
                                    PressedForeThemeColorIndex =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =10890
                                    Top =5340
                                    Width =1875
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="ExitOptionsConsole"
                                    Caption ="EXIT Options Dashboard"
                                    StatusBarText ="Enter SUF Print dt to print Surrender Unit Forms"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Find Next"
                                    LeftPadding =90
                                    RightPadding =105
                                    BottomPadding =165
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =10890
                                    LayoutCachedTop =5340
                                    LayoutCachedWidth =12765
                                    LayoutCachedHeight =5580
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =8872517
                                    OldBorderStyle =0
                                    BorderThemeColorIndex =4
                                    BorderShade =75.0
                                    HoverThemeColorIndex =2
                                    HoverTint =20.0
                                    PressedColor =8872517
                                    HoverForeColor =6108695
                                    PressedForeColor =0
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =14910
                                    Top =5340
                                    Width =1875
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="ExitOptionsHistory"
                                    Caption ="Group EXIT History"
                                    OnClick ="[Event Procedure]"
                                    Tag ="UAT"
                                    LeftPadding =90
                                    RightPadding =105
                                    BottomPadding =165
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =14910
                                    LayoutCachedTop =5340
                                    LayoutCachedWidth =16785
                                    LayoutCachedHeight =5580
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =9211020
                                    OldBorderStyle =0
                                    BorderThemeColorIndex =4
                                    BorderShade =75.0
                                    HoverColor =12566463
                                    PressedColor =9211020
                                    HoverForeThemeColorIndex =0
                                    PressedForeThemeColorIndex =1
                                    Overlaps =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =5160
                            Width =19305
                            Height =9030
                            Name ="TrackGroupChanges"
                            Caption ="Track Unit/Group Changes"
                            LayoutCachedLeft =135
                            LayoutCachedTop =5160
                            LayoutCachedWidth =19440
                            LayoutCachedHeight =14190
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =255
                                    Top =5340
                                    Width =13230
                                    Height =7695
                                    Name ="GroupRevisionLog"
                                    SourceObject ="Form.GroupRevisionLog"
                                    LinkChildFields ="GroupID"
                                    LinkMasterFields ="SGrpID"

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =5340
                                    LayoutCachedWidth =13485
                                    LayoutCachedHeight =13035
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =7
                                    Left =13665
                                    Top =5565
                                    Width =5550
                                    Height =4710
                                    FontSize =9
                                    TabIndex =1
                                    Name ="UnitRevDetail"
                                    RowSourceType ="Table/Query"
                                    ColumnWidths ="720;720;720;720;0;0;0"
                                    OnDblClick ="[Event Procedure]"

                                    LayoutCachedLeft =13665
                                    LayoutCachedTop =5565
                                    LayoutCachedWidth =19215
                                    LayoutCachedHeight =10275
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =13665
                                            Top =5340
                                            Width =5550
                                            Height =240
                                            FontSize =9
                                            ForeColor =8872517
                                            Name ="UnitRevDetailLabel"
                                            Caption ="Unit Detail"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =13665
                                            LayoutCachedTop =5340
                                            LayoutCachedWidth =19215
                                            LayoutCachedHeight =5580
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =255
                                    Left =12165
                                    Top =5925
                                    Width =870
                                    Height =315
                                    Name ="Label4375"
                                    Caption ="List4374:"
                                    LayoutCachedLeft =12165
                                    LayoutCachedTop =5925
                                    LayoutCachedWidth =13035
                                    LayoutCachedHeight =6240
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    CanGrow = NotDefault
                                    CanShrink = NotDefault
                                    SpecialEffect =5
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13665
                                    Top =11205
                                    Width =5550
                                    Height =1605
                                    FontSize =9
                                    TabIndex =2
                                    BorderColor =0
                                    ForeColor =6108695
                                    Name ="ReasonNote"
                                    StatusBarText ="1st option lrf"
                                    ValidationText ="Rent must be >=0"
                                    OnDblClick ="[Event Procedure]"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    TextFormat =1

                                    LayoutCachedLeft =13665
                                    LayoutCachedTop =11205
                                    LayoutCachedWidth =19215
                                    LayoutCachedHeight =12810
                                    RowStart =1
                                    RowEnd =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =13665
                                            Top =10965
                                            Width =5550
                                            Height =240
                                            FontSize =9
                                            ForeColor =8872517
                                            Name ="ReasonNoteLabel"
                                            Caption ="Action Reason / Note"
                                            FontName ="Segoe UI"
                                            LeftPadding =0
                                            TopPadding =0
                                            RightPadding =0
                                            BottomPadding =0
                                            LayoutCachedLeft =13665
                                            LayoutCachedTop =10965
                                            LayoutCachedWidth =19215
                                            LayoutCachedHeight =11205
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    CanGrow = NotDefault
                                    CanShrink = NotDefault
                                    SpecialEffect =5
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13665
                                    Top =10515
                                    Width =5550
                                    Height =420
                                    FontSize =9
                                    TabIndex =3
                                    BorderColor =0
                                    ForeColor =6108695
                                    Name ="ReasonAction"
                                    StatusBarText ="1st option lrf"
                                    DefaultValue ="0"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =13665
                                    LayoutCachedTop =10515
                                    LayoutCachedWidth =19215
                                    LayoutCachedHeight =10935
                                    RowStart =1
                                    RowEnd =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =13665
                                            Top =10275
                                            Width =5550
                                            Height =240
                                            FontSize =9
                                            ForeColor =8872517
                                            Name ="Label4385"
                                            Caption ="Action Desc"
                                            FontName ="Segoe UI"
                                            LeftPadding =0
                                            TopPadding =0
                                            RightPadding =0
                                            BottomPadding =0
                                            LayoutCachedLeft =13665
                                            LayoutCachedTop =10275
                                            LayoutCachedWidth =19215
                                            LayoutCachedHeight =10515
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    DecimalPlaces =4
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =18465
                                    Top =12885
                                    Width =705
                                    Height =210
                                    FontSize =8
                                    TabIndex =4
                                    Name ="ReasonINDX"
                                    Format ="Percent"
                                    StatusBarText ="int rt for int/pd bill"
                                    Tag ="StayOpen"

                                    LayoutCachedLeft =18465
                                    LayoutCachedTop =12885
                                    LayoutCachedWidth =19170
                                    LayoutCachedHeight =13095
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4839
                    Top =1410
                    Width =1320
                    Height =360
                    TabIndex =29
                    BackColor =-2147483617
                    BorderColor =-2147483617
                    ForeColor =-2147483607
                    Name ="PTExp"
                    ControlSource ="=IIf(IsNull([BLCD]),\"tbd\",DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1)"
                    Format ="Short Date"

                    LayoutCachedLeft =4839
                    LayoutCachedTop =1410
                    LayoutCachedWidth =6159
                    LayoutCachedHeight =1770
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10308
                    Top =3900
                    Width =4740
                    Height =255
                    FontSize =10
                    TabIndex =60
                    Name ="EarlyTermNote"
                    ControlSource ="EarlyTermNote"
                    GroupTable =49

                    LayoutCachedLeft =10308
                    LayoutCachedTop =3900
                    LayoutCachedWidth =15048
                    LayoutCachedHeight =4155
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9156
                            Top =3900
                            Width =1080
                            Height =255
                            FontSize =10
                            Name ="Label83"
                            Caption ="Early Note:"
                            GroupTable =49
                            LayoutCachedLeft =9156
                            LayoutCachedTop =3900
                            LayoutCachedWidth =10236
                            LayoutCachedHeight =4155
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =6
                            GroupTable =49
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =8532
                    Top =4188
                    Width =228
                    Height =252
                    TabIndex =63
                    BorderColor =12632256
                    Name ="BuyOutTBD"
                    ControlSource ="BuyOutTBD"
                    Tag ="Protected"
                    GroupTable =49
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8532
                    LayoutCachedTop =4188
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =4440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =8760
                            Top =4188
                            Width =360
                            Height =252
                            FontSize =10
                            Name ="Label84"
                            Caption ="TBD"
                            GroupTable =49
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =8760
                            LayoutCachedTop =4188
                            LayoutCachedWidth =9120
                            LayoutCachedHeight =4440
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =6
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =49
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6924
                    Top =4188
                    Width =1260
                    Height =252
                    FontSize =10
                    TabIndex =62
                    Name ="BuyOutPrice"
                    ControlSource ="BuyOutPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="BuyOutPrice must be >= )"
                    DefaultValue ="0"
                    GroupTable =49
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6924
                    LayoutCachedTop =4188
                    LayoutCachedWidth =8184
                    LayoutCachedHeight =4440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =5652
                            Top =4188
                            Width =1200
                            Height =252
                            FontSize =10
                            Name ="Label86"
                            Caption ="Buy Out Price:"
                            GroupTable =49
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =5652
                            LayoutCachedTop =4188
                            LayoutCachedWidth =6852
                            LayoutCachedHeight =4440
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =6
                            GroupTable =49
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    Left =8244
                    Top =4188
                    Width =252
                    Height =252
                    FontSize =10
                    Name ="Label87"
                    Caption ="or"
                    GroupTable =49
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8244
                    LayoutCachedTop =4188
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =49
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =17790
                    Top =540
                    Width =1560
                    Height =285
                    FontSize =8
                    TabIndex =23
                    Name ="SUF_Selection"
                    Caption ="SUF Selection Form"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17790
                    LayoutCachedTop =540
                    LayoutCachedWidth =19350
                    LayoutCachedHeight =825
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14474460
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =16260
                    Top =540
                    Width =1335
                    Height =1109
                    FontSize =8
                    Name ="OpenMassMod"
                    Caption ="Mass Mod\015\012Units - Child Parts - COA \015\012"
                    StatusBarText ="VINS must be entered, Status Must be B and CofA sent date must be nu.."
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =16260
                    LayoutCachedTop =540
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =1649
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14474460
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10308
                    Top =4188
                    Width =4740
                    Height =252
                    FontSize =10
                    TabIndex =64
                    Name ="BuyOutNote"
                    ControlSource ="BuyOutNote"
                    GroupTable =49
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10308
                    LayoutCachedTop =4188
                    LayoutCachedWidth =15048
                    LayoutCachedHeight =4440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9156
                            Top =4188
                            Width =1080
                            Height =252
                            FontSize =10
                            Name ="Label88"
                            Caption ="Buyout Note:"
                            GroupTable =49
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =9156
                            LayoutCachedTop =4188
                            LayoutCachedWidth =10236
                            LayoutCachedHeight =4440
                            RowStart =1
                            RowEnd =1
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =6
                            GroupTable =49
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =2
                    Left =8698
                    Top =1410
                    Width =1319
                    Height =360
                    TabIndex =32
                    Name ="LeasePaidAdvArr"
                    ControlSource ="LeasePaidAdvArr"
                    RowSourceType ="Value List"
                    RowSource ="\"Adv\";\"Arr\""
                    ColumnWidths ="720"
                    Tag ="Protected"

                    LayoutCachedLeft =8698
                    LayoutCachedTop =1410
                    LayoutCachedWidth =10017
                    LayoutCachedHeight =1770
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =7293
                    Top =1410
                    Width =1319
                    Height =360
                    TabIndex =31
                    Name ="LeasePymtFreq"
                    ControlSource ="LeasePymtFreq"
                    RowSourceType ="Value List"
                    RowSource ="M;MONTHLY;Q;QUARTERLY;S;SEMI-ANN;A;ANNUAL"
                    ColumnWidths ="0;720"
                    StatusBarText ="pmt period"
                    Tag ="Protected"

                    LayoutCachedLeft =7293
                    LayoutCachedTop =1410
                    LayoutCachedWidth =8612
                    LayoutCachedHeight =1770
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18072
                    Top =2364
                    Width =1296
                    Height =288
                    FontSize =9
                    TabIndex =20
                    BackColor =10092543
                    Name ="RALSent"
                    ControlSource ="SGrpRALO"
                    Format ="Short Date"
                    StatusBarText ="date RAL sent out"
                    GroupTable =60

                    LayoutCachedLeft =18072
                    LayoutCachedTop =2364
                    LayoutCachedWidth =19368
                    LayoutCachedHeight =2652
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =13
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =60
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =17340
                            Top =2364
                            Width =690
                            Height =288
                            FontSize =9
                            TopMargin =29
                            BackColor =10092543
                            Name ="Label422"
                            Caption =" Sent"
                            GroupTable =60
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =17340
                            LayoutCachedTop =2364
                            LayoutCachedWidth =18030
                            LayoutCachedHeight =2652
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =13
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =60
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14988
                    Top =1572
                    Width =816
                    FontSize =10
                    TabIndex =18
                    Name ="ExtendFSLPymt"
                    ControlSource ="ExtendFSLPymt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL ext pmts - non taxable"
                    ControlTipText ="\"This is in use for Toyota Only\""
                    GroupTable =57

                    LayoutCachedLeft =14988
                    LayoutCachedTop =1572
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =1812
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Top =1572
                    Width =1530
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="TslTxLBL"
                    ControlSource ="=\"Ext Tx Charge\""
                    ControlTipText ="\"This is in use for Toyota Only\""
                    GroupTable =57

                    LayoutCachedLeft =13380
                    LayoutCachedTop =1572
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =1812
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14988
                    Top =1260
                    Width =816
                    FontSize =10
                    TabIndex =16
                    Name ="SGrpExtFSLPmt"
                    ControlSource ="SGrpExtFSLPmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL ext pmts - non taxable"
                    ControlTipText ="\"This is in use for Toyota Only\""
                    GroupTable =57

                    LayoutCachedLeft =14988
                    LayoutCachedTop =1260
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Top =1260
                    Width =1530
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="EFSLLBL"
                    ControlSource ="=\"Ext Charge\""
                    ControlTipText ="\"This is in use for Toyota Only\""
                    GroupTable =57

                    LayoutCachedLeft =13380
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =1500
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Top =900
                    Width =1530
                    Height =288
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="AncillaryFeeLabel"
                    ControlSource ="=IIf([clientshnm]=\"toyota logistics\",\"Non-Serv Fee\",\"Ancillary Fee\")"
                    GroupTable =57

                    LayoutCachedLeft =13380
                    LayoutCachedTop =900
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =1188
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Top =540
                    Width =1530
                    Height =285
                    FontSize =10
                    TabIndex =11
                    BorderColor =0
                    Name ="ServFeeLabel"
                    ControlSource ="=IIf([clientshnm]=\"costco\",\"Full Serv Fee\",\"Service Fee\")"
                    GroupTable =57

                    LayoutCachedLeft =13380
                    LayoutCachedTop =540
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =825
                    LayoutGroup =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14988
                    Top =900
                    Width =816
                    Height =288
                    FontSize =10
                    TabIndex =14
                    Name ="SGrpRSLTPmt"
                    ControlSource ="SGrpRSLTPmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    Tag ="TrackUpdate"
                    GroupTable =57

                    LayoutCachedLeft =14988
                    LayoutCachedTop =900
                    LayoutCachedWidth =15804
                    LayoutCachedHeight =1188
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =57
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18072
                    Top =2724
                    Width =1296
                    Height =288
                    FontSize =9
                    TabIndex =21
                    BackColor =6737151
                    Name ="RALBack"
                    ControlSource ="SGrpRALB"
                    Format ="Short Date"
                    StatusBarText ="date RAL Returned"
                    GroupTable =60

                    LayoutCachedLeft =18072
                    LayoutCachedTop =2724
                    LayoutCachedWidth =19368
                    LayoutCachedHeight =3012
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =13
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =60
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =17340
                            Top =2724
                            Width =690
                            Height =288
                            FontSize =9
                            TopMargin =29
                            BackColor =6737151
                            Name ="Label339"
                            Caption ="Return"
                            GroupTable =60
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =17340
                            LayoutCachedTop =2724
                            LayoutCachedWidth =18030
                            LayoutCachedHeight =3012
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =13
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =60
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =360
                    Top =3465
                    Height =255
                    TabIndex =55
                    BorderColor =12632256
                    Name ="SGRPINTPD"
                    ControlSource ="SGRPINTPD"
                    StatusBarText ="int/pd bill"

                    LayoutCachedLeft =360
                    LayoutCachedTop =3465
                    LayoutCachedWidth =620
                    LayoutCachedHeight =3720
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =570
                            Top =3420
                            Width =1185
                            Height =270
                            FontSize =10
                            Name ="INtPDLBL"
                            Caption ="Interest/PDs"
                            LayoutCachedLeft =570
                            LayoutCachedTop =3420
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =3690
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =4
                    OverlapFlags =215
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =1140
                    Top =3750
                    Width =705
                    Height =270
                    FontSize =9
                    TabIndex =58
                    Name ="SGRPINTRT"
                    ControlSource ="SGRPINTRT"
                    Format ="Percent"
                    StatusBarText ="int rt for int/pd bill"
                    Tag ="StayOpen"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =3750
                    LayoutCachedWidth =1845
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =204
                            Left =270
                            Top =3765
                            Width =780
                            Height =255
                            FontSize =9
                            Name ="INtPDRTLBL"
                            Caption ="Int Rate:"
                            LayoutCachedLeft =270
                            LayoutCachedTop =3765
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =4020
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15135
                    Top =165
                    Width =720
                    FontSize =10
                    TabIndex =1
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="FeeTerm"
                    ControlSource ="ScheduleGroupPaymentTerm"
                    Format ="General Number"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="StayOpen"

                    LayoutCachedLeft =15135
                    LayoutCachedTop =165
                    LayoutCachedWidth =15855
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13680
                            Top =75
                            Width =1440
                            Height =330
                            FontSize =10
                            FontWeight =700
                            Name ="Label486"
                            Caption ="Fee Term Mths"
                            LayoutCachedLeft =13680
                            LayoutCachedTop =75
                            LayoutCachedWidth =15120
                            LayoutCachedHeight =405
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =3360
                    Width =3285
                    FontSize =9
                    TabIndex =53
                    Name ="exdeflbl"
                    ControlSource ="=IIf(IsNull([BLCD]) Or [ExchangeITMonth]=0,Null,\"earliest new unit in service \""
                        " & DateAdd(\"m\",[ExchangeITMonth],[BLCD]))"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =3360
                    LayoutCachedWidth =12765
                    LayoutCachedHeight =3600
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =13080
                    Top =60
                    Width =600
                    Height =360
                    FontSize =14
                    FontWeight =700
                    ForeColor =5330263
                    Name ="Label489"
                    Caption ="FEES"
                    LayoutCachedLeft =13080
                    LayoutCachedTop =60
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =420
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7920
                    Left =255
                    Top =300
                    Width =1620
                    Height =330
                    TabIndex =2
                    Name ="SGrpFAID"
                    ControlSource ="SGrpFAID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Trim([Deals].[faid]) AS FAID, Deals.dealShortDesc, Clients.clientID FROM "
                        "Clients LEFT JOIN Deals ON Clients.clientID = Deals.clientRef WHERE (((Clients.c"
                        "lientID)=[Forms]![GroupForm]![clientID])) ORDER BY Trim([Deals].[faid]) DESC;"
                    ColumnWidths ="1440;4752"
                    StatusBarText ="Enter ONLY if different from schedule FAID"
                    Tag ="Protected"
                    ControlTipText ="Entere ONLY if different from schedule FAID"
                    AllowValueListEdits =0

                    LayoutCachedLeft =255
                    LayoutCachedTop =300
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =6
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6408
                    Top =3336
                    Width =1056
                    Height =288
                    FontSize =9
                    TabIndex =49
                    BorderColor =10855845
                    Name ="MTMRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Must be greater than or equal to 0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    GroupTable =13

                    LayoutCachedLeft =6408
                    LayoutCachedTop =3336
                    LayoutCachedWidth =7464
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =5784
                            Top =3336
                            Width =552
                            Height =288
                            FontSize =10
                            Name ="Label2525"
                            Caption ="LRF%"
                            GroupTable =13
                            LayoutCachedLeft =5784
                            LayoutCachedTop =3336
                            LayoutCachedWidth =6336
                            LayoutCachedHeight =3624
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =4050
                    Top =3150
                    Width =240
                    Height =225
                    FontSize =10
                    Name ="Label497"
                    Caption ="┤"
                    LayoutCachedLeft =4050
                    LayoutCachedTop =3150
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =3375
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =19080
                    Top =60
                    Width =240
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="AddIn Fees"
                    Caption ="≡"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="AddIn_Fees"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19080
                    LayoutCachedTop =60
                    LayoutCachedWidth =19320
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    BorderThemeColorIndex =1
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =13845
                    Top =240
                    Width =1215
                    Height =210
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label488"
                    Caption ="0 = no term set"
                    LayoutCachedLeft =13845
                    LayoutCachedTop =240
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =450
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =4065
                    Top =3000
                    Width =600
                    Height =240
                    FontSize =10
                    Name ="Label499"
                    Caption ="┌────"
                    LayoutCachedLeft =4065
                    LayoutCachedTop =3000
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =3240
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4908
                    Top =2988
                    Width =816
                    Height =276
                    FontSize =10
                    TabIndex =46
                    Name ="LeaseExtOption1Term"
                    ControlSource ="LeaseExtOption1Term"
                    StatusBarText ="1st option in months"
                    ControlTipText ="1st option in months"
                    GroupTable =13

                    LayoutCachedLeft =4908
                    LayoutCachedTop =2988
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =3264
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =13
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =4248
                            Top =2988
                            Width =588
                            Height =276
                            FontSize =10
                            Name ="Label2516"
                            Caption ="Term"
                            GroupTable =13
                            LayoutCachedLeft =4248
                            LayoutCachedTop =2988
                            LayoutCachedWidth =4836
                            LayoutCachedHeight =3264
                            LayoutGroup =2
                            BackThemeColorIndex =1
                            BackShade =85.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4908
                    Top =3336
                    Width =816
                    Height =288
                    FontSize =10
                    BorderColor =12632256
                    Name ="Label438"
                    Caption ="M-T-M"
                    GroupTable =13
                    LayoutCachedLeft =4908
                    LayoutCachedTop =3336
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =13
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =10155
                    Top =1410
                    Width =1365
                    Height =360
                    TabIndex =33
                    BoundColumn =1
                    Name ="AssignedIndicator"
                    RowSourceType ="Value List"
                    RowSource ="Final (Synd);F;FA Hold;H;FA Inv;I;Projected;P"
                    ColumnWidths ="2448;288"
                    ValidationRule ="=\"F\" Or \"H\" Or \"I\" Or \"P\""
                    ValidationText ="Must be H, I, F or P"
                    DefaultValue ="\"P\""
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Protected"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =10155
                    LayoutCachedTop =1410
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1770
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4248
                    Top =3336
                    Width =588
                    Height =288
                    FontSize =10
                    Name ="Label2514"
                    Caption ="Term"
                    GroupTable =13
                    LayoutCachedLeft =4248
                    LayoutCachedTop =3336
                    LayoutCachedWidth =4836
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =13
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =2040
                    Top =2820
                    Width =1200
                    Height =900
                    FontSize =10
                    FontWeight =700
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3990"
                    Caption ="Extend IT"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =3720
                    BackThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =2040
                    Top =3720
                    Width =1200
                    Height =900
                    FontSize =10
                    FontWeight =700
                    BackColor =3342489
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label3991"
                    Caption ="Terminate"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =3720
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4620
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3255
                    Top =4005
                    Width =780
                    Height =255
                    FontSize =10
                    FontWeight =700
                    Name ="Label3993"
                    Caption ="Options"
                    LayoutCachedLeft =3255
                    LayoutCachedTop =4005
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =4260
                End
                Begin EmptyCell
                    Left =6924
                    Top =3900
                    Width =1260
                    Height =255
                    Name ="EmptyCell4010"
                    GroupTable =49
                    LayoutCachedLeft =6924
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8184
                    LayoutCachedHeight =4155
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =49
                End
                Begin EmptyCell
                    Left =8244
                    Top =3900
                    Width =252
                    Height =255
                    Name ="EmptyCell4015"
                    GroupTable =49
                    LayoutCachedLeft =8244
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4155
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =6
                    GroupTable =49
                End
                Begin EmptyCell
                    Left =8532
                    Top =3900
                    Width =228
                    Height =255
                    Name ="EmptyCell4019"
                    GroupTable =49
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8532
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =4155
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =6
                    GroupTable =49
                End
                Begin EmptyCell
                    Left =8760
                    Top =3900
                    Width =360
                    Height =255
                    Name ="EmptyCell4023"
                    GroupTable =49
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8760
                    LayoutCachedTop =3900
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =4155
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =6
                    GroupTable =49
                End
                Begin ComboBox
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3300
                    Top =2280
                    Width =840
                    FontSize =10
                    FontWeight =700
                    TabIndex =38
                    Name ="ExchangeITGroup"
                    ControlSource ="ExchangeITGroup"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;576"
                    Tag ="Protected"
                    Format ="True/False"
                    BottomPadding =150

                    LayoutCachedLeft =3300
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =2520
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    Left =2040
                    Top =2040
                    Width =1200
                    Height =780
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483600
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Label4042"
                    Caption ="Exchange IT"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =2040
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =2820
                End
                Begin Label
                    OverlapFlags =223
                    Left =4095
                    Top =4005
                    Width =240
                    Height =225
                    FontSize =10
                    Name ="Label4059"
                    Caption ="┤"
                    LayoutCachedLeft =4095
                    LayoutCachedTop =4005
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =4230
                End
                Begin Label
                    OverlapFlags =223
                    Left =4110
                    Top =3855
                    Width =600
                    Height =240
                    FontSize =10
                    Name ="Label4060"
                    Caption ="┌────"
                    LayoutCachedLeft =4110
                    LayoutCachedTop =3855
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =4095
                End
                Begin Label
                    OverlapFlags =223
                    Left =4110
                    Top =4185
                    Width =600
                    Height =240
                    FontSize =10
                    Name ="Label4058"
                    Caption ="└────"
                    LayoutCachedLeft =4110
                    LayoutCachedTop =4185
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =4425
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10155
                    Top =1200
                    Width =1365
                    Height =210
                    FontSize =9
                    Name ="Label4069"
                    Caption ="Status"
                    LayoutCachedLeft =10155
                    LayoutCachedTop =1200
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1410
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11595
                    Top =1200
                    Width =1260
                    Height =240
                    FontSize =10
                    Name ="BLCDLBL"
                    Caption ="Date Assigned"
                    LayoutCachedLeft =11595
                    LayoutCachedTop =1200
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =1440
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4830
                    Top =1200
                    Width =1335
                    Height =240
                    FontSize =10
                    BackColor =14211288
                    Name ="Label4111"
                    Caption ="Primary Expiry"
                    LayoutCachedLeft =4830
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =1440
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =240
                    Top =2280
                    Width =1740
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =37
                    Name ="PerDiemStartsOn"
                    ControlSource ="PerDiemStartsOn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.StartsOn;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"

                    LayoutCachedLeft =240
                    LayoutCachedTop =2280
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =2610
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =105
                            Top =2040
                            Width =1875
                            Height =240
                            FontSize =10
                            Name ="Label4113"
                            Caption ="Per Diem Starts On"
                            LayoutCachedLeft =105
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1980
                            LayoutCachedHeight =2280
                            BackThemeColorIndex =9
                            BackTint =60.0
                            BorderThemeColorIndex =1
                            BorderShade =50.0
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =3615
                    Top =1740
                    Width =1125
                    Height =210
                    FontSize =9
                    Name ="Label4114"
                    Caption ="Committed"
                    LayoutCachedLeft =3615
                    LayoutCachedTop =1740
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =1950
                    ForeThemeColorIndex =2
                End
                Begin ComboBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =2160
                    Left =240
                    Top =1410
                    Width =1740
                    Height =330
                    ColumnOrder =14
                    FontSize =10
                    FontWeight =700
                    TabIndex =26
                    Name ="SgrpCommInd"
                    ControlSource ="SgrpCommInd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_LeaseStartsOn.ID, Lookup_LeaseStartsOn.StartsOn FROM Lookup_LeaseS"
                        "tartsOn ORDER BY Lookup_LeaseStartsOn.ID;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="blcd source accepted or in-service"

                    LayoutCachedLeft =240
                    LayoutCachedTop =1410
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =1740
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6330
                    Top =1200
                    Width =765
                    Height =240
                    FontSize =10
                    BackColor =14211288
                    Name ="Label4116"
                    Caption ="Notice"
                    LayoutCachedLeft =6330
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =1440
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =6240
                    Top =1740
                    Width =1020
                    Height =225
                    FontSize =9
                    Name ="Label4117"
                    Caption ="Months Req."
                    LayoutCachedLeft =6240
                    LayoutCachedTop =1740
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1965
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2176
                    Top =1200
                    Width =1320
                    Height =210
                    FontSize =10
                    BackColor =14211288
                    Name ="Label4118"
                    Caption ="BLCD"
                    LayoutCachedLeft =2176
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3496
                    LayoutCachedHeight =1410
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =223
                    TextAlign =3
                    Left =5415
                    Top =660
                    Width =165
                    Height =240
                    FontSize =9
                    Name ="Label4143"
                    Caption ="or"
                    LayoutCachedLeft =5415
                    LayoutCachedTop =660
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3525
                    Top =90
                    Width =750
                    Height =210
                    FontSize =10
                    Name ="Nofounitslbl"
                    Caption ="Units In"
                    LayoutCachedLeft =3525
                    LayoutCachedTop =90
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7293
                    Top =1200
                    Width =1319
                    Height =210
                    FontSize =10
                    Name ="PisLBL"
                    Caption ="Payments are"
                    LayoutCachedLeft =7293
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8612
                    LayoutCachedHeight =1410
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8698
                    Top =1200
                    Width =1319
                    Height =210
                    FontSize =10
                    Name ="MdeinLBL"
                    Caption ="Paid in"
                    LayoutCachedLeft =8698
                    LayoutCachedTop =1200
                    LayoutCachedWidth =10017
                    LayoutCachedHeight =1410
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =10110
                    Top =90
                    Width =2805
                    Height =240
                    FontSize =10
                    BackColor =26112
                    Name ="Label4158"
                    Caption ="Assignment "
                    LayoutCachedLeft =10110
                    LayoutCachedTop =90
                    LayoutCachedWidth =12915
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =5596
                    Top =90
                    Width =1499
                    Height =210
                    FontSize =10
                    Name ="grpoeclbl"
                    Caption ="Group OEC"
                    LayoutCachedLeft =5596
                    LayoutCachedTop =90
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =7202
                    Top =90
                    Width =1410
                    Height =210
                    FontSize =10
                    Name ="LPMTLBL"
                    Caption ="Lease Payment"
                    LayoutCachedLeft =7202
                    LayoutCachedTop =90
                    LayoutCachedWidth =8612
                    LayoutCachedHeight =300
                    RowStart =3
                    RowEnd =3
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    Left =7184
                    Top =630
                    Width =1410
                    Height =270
                    FontSize =9
                    Name ="Label4253"
                    Caption ="for Group"
                    LayoutCachedLeft =7184
                    LayoutCachedTop =630
                    LayoutCachedWidth =8594
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =8698
                    Top =90
                    Width =1320
                    Height =210
                    FontSize =10
                    Name ="LRFLBL"
                    Caption ="LRF"
                    LayoutCachedLeft =8698
                    LayoutCachedTop =90
                    LayoutCachedWidth =10018
                    LayoutCachedHeight =300
                    RowStart =3
                    RowEnd =3
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =3525
                    Top =630
                    Width =750
                    Height =270
                    FontSize =9
                    Name ="Label4257"
                    Caption ="Group"
                    LayoutCachedLeft =3525
                    LayoutCachedTop =630
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =2
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =5460
                    Top =900
                    Width =1815
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="Label4258"
                    Caption ="Buyout if Capital Lease"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =900
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =1140
                    ForeThemeColorIndex =0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =2
                    ListWidth =720
                    Left =4455
                    Top =300
                    Width =900
                    Height =330
                    TabIndex =5
                    Name ="SGRPPLB"
                    ControlSource ="SGRPPLB"
                    RowSourceType ="Value List"
                    RowSource ="Yes;No"
                    ColumnWidths ="720"
                    DefaultValue ="No"
                    Tag ="Protected"
                    ControlTipText ="Is this a Purchase Lease Back (PLB)?"
                    Format ="Yes/No"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4455
                    LayoutCachedTop =300
                    LayoutCachedWidth =5355
                    LayoutCachedHeight =630
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4500
                    Top =90
                    Width =810
                    Height =210
                    FontSize =10
                    Name ="Label4261"
                    Caption ="SLB (PLB)"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =90
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =50.0
                    ForeThemeColorIndex =1
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =4272
                    Top =3900
                    Width =180
                    Height =255
                    TabIndex =59
                    BorderColor =12632256
                    Name ="EarlyTermOptional\015\012"
                    ControlSource ="EarlyTermOptional"
                    Tag ="Protected"
                    EventProcPrefix ="EarlyTermOptional__"
                    GroupTable =49

                    LayoutCachedLeft =4272
                    LayoutCachedTop =3900
                    LayoutCachedWidth =4452
                    LayoutCachedHeight =4155
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =4512
                            Top =3900
                            Width =2340
                            Height =255
                            FontSize =10
                            Name ="Label82"
                            Caption ="Early Termination"
                            GroupTable =49
                            LayoutCachedLeft =4512
                            LayoutCachedTop =3900
                            LayoutCachedWidth =6852
                            LayoutCachedHeight =4155
                            ColumnStart =1
                            ColumnEnd =2
                            LayoutGroup =6
                            BackThemeColorIndex =1
                            BackShade =95.0
                            GroupTable =49
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =4272
                    Top =4188
                    Width =180
                    Height =252
                    TabIndex =61
                    BorderColor =12632256
                    Name ="FMVBuyOutOptional"
                    ControlSource ="FMVBuyOutOptional"
                    StatusBarText ="FMV Buyout at Expiry?"
                    Tag ="Protected"
                    GroupTable =49
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4272
                    LayoutCachedTop =4188
                    LayoutCachedWidth =4452
                    LayoutCachedHeight =4440
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =6
                    GroupTable =49
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =4512
                            Top =4188
                            Width =1080
                            Height =252
                            FontSize =10
                            Name ="Label85"
                            Caption ="FMV Buyout"
                            GroupTable =49
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4512
                            LayoutCachedTop =4188
                            LayoutCachedWidth =5592
                            LayoutCachedHeight =4440
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =6
                            BackThemeColorIndex =1
                            BackShade =95.0
                            GroupTable =49
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =223
                    Left =60
                    Top =1140
                    Width =115
                    Height =900
                    BackColor =26112
                    Name ="Box4284"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1140
                    LayoutCachedWidth =175
                    LayoutCachedHeight =2040
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =215
                    Left =60
                    Top =2040
                    Width =115
                    Height =2592
                    BackColor =5532596
                    BorderColor =5532596
                    Name ="Box4286"
                    LayoutCachedLeft =60
                    LayoutCachedTop =2040
                    LayoutCachedWidth =175
                    LayoutCachedHeight =4632
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =210
                    Top =1200
                    Width =1785
                    Height =240
                    FontSize =10
                    BackColor =14211288
                    Name ="CommBaselbl"
                    Caption ="BLCD Basis"
                    LayoutCachedLeft =210
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =1440
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =18090
                    Top =3480
                    Width =1260
                    Height =240
                    FontSize =10
                    Name ="Label143"
                    Caption ="Bank Sch #"
                    LayoutCachedLeft =18090
                    LayoutCachedTop =3480
                    LayoutCachedWidth =19350
                    LayoutCachedHeight =3720
                    ForeThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =13020
                    Top =2400
                    Width =540
                    Height =240
                    FontSize =10
                    ForeColor =6108695
                    Name ="Label4327"
                    Caption ="Group"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13020
                    LayoutCachedTop =2400
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =2640
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =2700
                    Width =1299
                    FontSize =9
                    TabIndex =43
                    BorderColor =14270637
                    Name ="CalcSalesRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14535
                    LayoutCachedTop =2700
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =2940
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Left =13845
                            Top =2700
                            Width =660
                            Height =240
                            FontSize =10
                            ForeColor =6108695
                            Name ="Label3589"
                            Caption ="LO Rent"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =13845
                            LayoutCachedTop =2700
                            LayoutCachedWidth =14505
                            LayoutCachedHeight =2940
                            RowStart =3
                            RowEnd =3
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =3180
                    Width =1299
                    FontSize =9
                    TabIndex =52
                    BorderColor =14270637
                    ForeColor =8210719
                    Name ="CalcRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14535
                    LayoutCachedTop =3180
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =3420
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Left =13680
                            Top =3180
                            Width =825
                            Height =240
                            FontSize =10
                            BorderColor =8210719
                            ForeColor =6108695
                            Name ="Label538"
                            Caption ="Synd Rent"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =13680
                            LayoutCachedTop =3180
                            LayoutCachedWidth =14505
                            LayoutCachedHeight =3420
                            RowStart =1
                            RowEnd =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =2220
                    Width =1299
                    FontSize =9
                    TabIndex =36
                    BackColor =15858167
                    BorderColor =14270637
                    ForeColor =8210719
                    Name ="CalcOEC"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14535
                    LayoutCachedTop =2220
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =2460
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Left =14145
                            Top =2220
                            Width =345
                            Height =240
                            FontSize =10
                            BorderColor =8210719
                            ForeColor =6108695
                            Name ="Label536"
                            Caption ="OEC"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =14145
                            LayoutCachedTop =2220
                            LayoutCachedWidth =14490
                            LayoutCachedHeight =2460
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8064
                    Top =2988
                    Width =1308
                    Height =276
                    FontSize =9
                    TabIndex =48
                    BorderColor =10855845
                    Name ="ExtRollUp"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="UAT"
                    GroupTable =13

                    LayoutCachedLeft =8064
                    LayoutCachedTop =2988
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =3264
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =13
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =7524
                            Top =2988
                            Width =468
                            Height =276
                            FontSize =10
                            Name ="Label2544"
                            Caption ="Rent"
                            GroupTable =13
                            LayoutCachedLeft =7524
                            LayoutCachedTop =2988
                            LayoutCachedWidth =7992
                            LayoutCachedHeight =3264
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =17790
                    Top =945
                    Width =1560
                    Height =285
                    FontSize =8
                    TabIndex =25
                    Name ="CloneSchGroup"
                    Caption ="Clone Sch Group"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17790
                    LayoutCachedTop =945
                    LayoutCachedWidth =19350
                    LayoutCachedHeight =1230
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10856415
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =17790
                    Top =1350
                    Width =1560
                    Height =285
                    FontSize =8
                    TabIndex =35
                    Name ="CopyThisGroup"
                    Caption ="Copy This Group"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Find Next"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17790
                    LayoutCachedTop =1350
                    LayoutCachedWidth =19350
                    LayoutCachedHeight =1635
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14474460
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =32
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8064
                    Top =3336
                    Width =1308
                    Height =288
                    FontSize =9
                    TabIndex =50
                    BorderColor =10855845
                    Name ="MTMRollup"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="UAT"
                    GroupTable =13

                    LayoutCachedLeft =8064
                    LayoutCachedTop =3336
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =13
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =7524
                            Top =3336
                            Width =468
                            Height =288
                            FontSize =10
                            Name ="Label2540"
                            Caption ="Rent"
                            GroupTable =13
                            LayoutCachedLeft =7524
                            LayoutCachedTop =3336
                            LayoutCachedWidth =7992
                            LayoutCachedHeight =3624
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            BackThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11505
                    Top =2280
                    Width =1230
                    FontSize =9
                    TabIndex =41
                    BorderColor =10855845
                    Name ="CanExchangeDAte"
                    ControlSource ="=IIf(Nz([ExchangeITMonth],0)>0 And IsDate([blcd]),DateAdd(\"m\",[ExchangeITMonth"
                        "],[BLCD]),Null)"
                    Format ="Short Date"
                    Tag ="UAT"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11505
                    LayoutCachedTop =2280
                    LayoutCachedWidth =12735
                    LayoutCachedHeight =2520
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =10425
                            Top =2280
                            Width =1035
                            Height =240
                            FontSize =8
                            ForeColor =6108695
                            Name ="Label4453"
                            Caption ="1st date avail"
                            Tag ="UAT"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10425
                            LayoutCachedTop =2280
                            LayoutCachedWidth =11460
                            LayoutCachedHeight =2520
                            RowStart =3
                            RowEnd =3
                            BackThemeColorIndex =1
                            BackShade =95.0
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10200
                    Top =922
                    Width =2640
                    Height =225
                    FontSize =8
                    TabIndex =24
                    Name ="Validate"
                    Caption ="Validate This Group"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
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

                    LayoutCachedLeft =10200
                    LayoutCachedTop =922
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =1147
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =2
                    BackThemeColorIndex =8
                    BackTint =60.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =3420
                    Width =1299
                    FontSize =9
                    TabIndex =54
                    BackColor =15858167
                    BorderColor =14270637
                    ForeColor =8210719
                    Name ="CalcSyndRV"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14535
                    LayoutCachedTop =3420
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =3660
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =13845
                            Top =3420
                            Width =660
                            Height =240
                            FontSize =10
                            BorderColor =8210719
                            ForeColor =6108695
                            Name ="Label4502"
                            Caption =" Synd RV"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =13845
                            LayoutCachedTop =3420
                            LayoutCachedWidth =14505
                            LayoutCachedHeight =3660
                            RowStart =2
                            RowEnd =2
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14535
                    Top =2460
                    Width =1299
                    FontSize =9
                    TabIndex =42
                    BackColor =14811130
                    BorderColor =14270637
                    ForeColor =8210719
                    Name ="CalcVendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14535
                    LayoutCachedTop =2460
                    LayoutCachedWidth =15834
                    LayoutCachedHeight =2700
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Left =13860
                            Top =2460
                            Width =675
                            Height =240
                            FontSize =10
                            BackColor =14811130
                            BorderColor =8210719
                            ForeColor =6108695
                            Name ="Label4524"
                            Caption ="Vendor"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =13860
                            LayoutCachedTop =2460
                            LayoutCachedWidth =14535
                            LayoutCachedHeight =2700
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18072
                    Top =2040
                    Width =1296
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    Name ="OpsPackage"
                    ControlSource ="OpsPackage"
                    Format ="Short Date"
                    StatusBarText ="date RAL sent out"
                    GroupTable =60
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18072
                    LayoutCachedTop =2040
                    LayoutCachedWidth =19368
                    LayoutCachedHeight =2328
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =13
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =60
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =17340
                            Top =2040
                            Width =690
                            Height =288
                            FontSize =9
                            TopMargin =29
                            Name ="Label4566"
                            Caption ="Done"
                            GroupTable =60
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =17340
                            LayoutCachedTop =2040
                            LayoutCachedWidth =18030
                            LayoutCachedHeight =2328
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =13
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =60
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =16260
                    Top =2040
                    Width =1080
                    Height =288
                    FontSize =9
                    TopMargin =29
                    Name ="Label4550"
                    Caption ="OPS Package"
                    GroupTable =60
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16260
                    LayoutCachedTop =2040
                    LayoutCachedWidth =17340
                    LayoutCachedHeight =2328
                    LayoutGroup =13
                    BackThemeColorIndex =4
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =60
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
' See "GroupForm.cls"
