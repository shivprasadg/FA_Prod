Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =1206
    Left =4680
    Top =2715
    Right =22410
    Bottom =18090
    Filter ="SchID=1834"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x4aecc6ab9638e540
    End
    RecordSource ="SELECT Clients.clientID, Clients.ClientGroupId, Schedule.*, MstrLease.MLNo, Clie"
        "nts.ClientShNm, MstrLease.MLDt, MstrLease.MLOrig, RALOutSchedule([Schid]) AS RAL"
        "Out, RALBackSchedule([Schid]) AS RALBack FROM (Clients INNER JOIN MstrLease ON C"
        "lients.clientID = MstrLease.Client) INNER JOIN Schedule ON MstrLease.MLID = Sche"
        "dule.SchMLRef;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
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
            BorderColor =12632256
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
            CanGrow = NotDefault
            Height =945
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =750
                    Width =18000
                    Height =195
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =750
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =945
                    TabIndex =6
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =11445
                    Top =210
                    Width =1335
                    Height =255
                    FontSize =10
                    TabIndex =1
                    Name ="ViewGrpBtn"
                    Caption ="View Groups"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11445
                    LayoutCachedTop =210
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =12840
                    Top =210
                    Width =1335
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="AddGroup"
                    Caption ="Add Group"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12840
                    LayoutCachedTop =210
                    LayoutCachedWidth =14175
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14220
                    Top =210
                    Width =1335
                    Height =255
                    FontSize =10
                    TabIndex =3
                    Name ="OpenClientsForm"
                    Caption ="Client Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14220
                    LayoutCachedTop =210
                    LayoutCachedWidth =15555
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =435
                    Width =11070
                    Height =465
                    ColumnOrder =0
                    FontSize =16
                    FontWeight =700
                    Name ="SchDesc"
                    ControlSource ="=UCase([clientshnm] & \" Schedule \" & [schNO] & IIf(IsNull([schalias]),Null,\"("
                        "\" & [schalias] & \")\"))"

                    LayoutCachedLeft =30
                    LayoutCachedTop =435
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =900
                    ThemeFontIndex =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =96
                    Width =2436
                    Height =525
                    FontSize =22
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Schedules"
                    FontName ="Segoe UI"
                    GroupTable =62
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =96
                    LayoutCachedWidth =2532
                    LayoutCachedHeight =525
                    LayoutGroup =10
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =62
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =16575
                    Top =255
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =4
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

                    LayoutCachedLeft =16575
                    LayoutCachedTop =255
                    LayoutCachedWidth =17130
                    LayoutCachedHeight =483
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
                    Left =17142
                    Top =255
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =17142
                    LayoutCachedTop =255
                    LayoutCachedWidth =17892
                    LayoutCachedHeight =483
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
            Height =14445
            Name ="Detail"
            Begin
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =5040
                    Left =5652
                    Top =468
                    Width =1992
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =967423
                    Name ="FSLType"
                    ControlSource ="FSLType"
                    RowSourceType ="Value List"
                    RowSource ="\"No\";\"Yes\";\"Yes-Include\";\"Yes-Perdiem\";\"Yes-Perdiem-Include\";\"Yes-Per"
                        "diem-Include-Refund\""
                    ColumnWidths ="1152"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="FAID is required"
                    GroupTable =2

                    LayoutCachedLeft =5652
                    LayoutCachedTop =468
                    LayoutCachedWidth =7644
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =4872
                            Top =468
                            Width =708
                            Height =252
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label168"
                            Caption ="FSL Type:"
                            GroupTable =2
                            LayoutCachedLeft =4872
                            LayoutCachedTop =468
                            LayoutCachedWidth =5580
                            LayoutCachedHeight =720
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =14370
                    Top =150
                    Width =0
                    Height =225
                    TabIndex =10
                    Name ="SchMLRef"
                    ControlSource ="SchMLRef"
                    StatusBarText ="reference to master lease"

                    LayoutCachedLeft =14370
                    LayoutCachedTop =150
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2784
                    Top =1812
                    Width =1068
                    Height =276
                    FontSize =9
                    TabIndex =16
                    Name ="SchOut"
                    ControlSource ="SchOut"
                    Format ="Short Date"
                    StatusBarText ="date schedule sent out"
                    GroupTable =19

                    LayoutCachedLeft =2784
                    LayoutCachedTop =1812
                    LayoutCachedWidth =3852
                    LayoutCachedHeight =2088
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1884
                            Top =1812
                            Width =828
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="Label5"
                            Caption ="Sch Out"
                            GroupTable =19
                            LayoutCachedLeft =1884
                            LayoutCachedTop =1812
                            LayoutCachedWidth =2712
                            LayoutCachedHeight =2088
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5244
                    Top =1812
                    Width =1068
                    Height =276
                    FontSize =9
                    TabIndex =17
                    Name ="SchIn"
                    ControlSource ="SchIn"
                    Format ="Short Date"
                    StatusBarText ="date schedule in"
                    GroupTable =19

                    LayoutCachedLeft =5244
                    LayoutCachedTop =1812
                    LayoutCachedWidth =6312
                    LayoutCachedHeight =2088
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3912
                            Top =1812
                            Width =1260
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="Label6"
                            Caption ="Sch IN"
                            GroupTable =19
                            LayoutCachedLeft =3912
                            LayoutCachedTop =1812
                            LayoutCachedWidth =5172
                            LayoutCachedHeight =2088
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =3
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8124
                    Top =1488
                    Width =252
                    TabIndex =19
                    Name ="SchTitleAssignee"
                    ControlSource ="SchTitleAssignee"
                    GroupTable =23

                    LayoutCachedLeft =8124
                    LayoutCachedTop =1488
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =1728
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6696
                            Top =1488
                            Width =1368
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label11"
                            Caption ="Assignee:"
                            GroupTable =23
                            LayoutCachedLeft =6696
                            LayoutCachedTop =1488
                            LayoutCachedWidth =8064
                            LayoutCachedHeight =1728
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8124
                    Top =1788
                    Width =252
                    TabIndex =21
                    Name ="SchTitltoFA"
                    ControlSource ="SchTitltoFA"
                    GroupTable =23

                    LayoutCachedLeft =8124
                    LayoutCachedTop =1788
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =2028
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6696
                            Top =1788
                            Width =1368
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label14"
                            Caption ="Fleet Advantage:"
                            GroupTable =23
                            LayoutCachedLeft =6696
                            LayoutCachedTop =1788
                            LayoutCachedWidth =8064
                            LayoutCachedHeight =2028
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8124
                    Top =2088
                    Width =252
                    TabIndex =22
                    Name ="SchTitletoClient"
                    ControlSource ="SchTitletoClient"
                    GroupTable =23

                    LayoutCachedLeft =8124
                    LayoutCachedTop =2088
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =2328
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6696
                            Top =2088
                            Width =1368
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Client:"
                            GroupTable =23
                            LayoutCachedLeft =6696
                            LayoutCachedTop =2088
                            LayoutCachedWidth =8064
                            LayoutCachedHeight =2328
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8124
                    Top =2700
                    Width =252
                    TabIndex =25
                    Name ="LienHolder"
                    ControlSource ="LienHolder"
                    GroupTable =23

                    LayoutCachedLeft =8124
                    LayoutCachedTop =2700
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =2940
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6696
                            Top =2700
                            Width =1368
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label16"
                            Caption ="Lien Holder"
                            GroupTable =23
                            LayoutCachedLeft =6696
                            LayoutCachedTop =2700
                            LayoutCachedWidth =8064
                            LayoutCachedHeight =2940
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9516
                    Top =2400
                    Width =4248
                    FontSize =9
                    TabIndex =24
                    Name ="LienHolderNm"
                    ControlSource ="LienHolderNm"
                    GroupTable =23

                    LayoutCachedLeft =9516
                    LayoutCachedTop =2400
                    LayoutCachedWidth =13764
                    LayoutCachedHeight =2640
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =8436
                            Top =2400
                            Width =1008
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Name:"
                            GroupTable =23
                            LayoutCachedLeft =8436
                            LayoutCachedTop =2400
                            LayoutCachedWidth =9444
                            LayoutCachedHeight =2640
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4092
                    Top =4116
                    Width =1176
                    Height =300
                    FontSize =9
                    TabIndex =36
                    Name ="SCHINSAMT"
                    ControlSource ="SCHINSAMT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="INSURANCE IS AT THE SCHEDULE LEVEL IF MASTER LEASE IS NOT BLANKET"
                    ValidationRule =">=0"
                    ValidationText ="Laibility Amt must be >=0"
                    DefaultValue ="0"
                    GroupTable =31

                    LayoutCachedLeft =4092
                    LayoutCachedTop =4116
                    LayoutCachedWidth =5268
                    LayoutCachedHeight =4416
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =31
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2364
                            Top =4116
                            Width =1656
                            Height =300
                            FontSize =9
                            FontWeight =700
                            Name ="Label19"
                            Caption ="Liability Amt:"
                            GroupTable =31
                            LayoutCachedLeft =2364
                            LayoutCachedTop =4116
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =4416
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =2
                            LayoutGroup =5
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =31
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2364
                    Top =3408
                    Width =936
                    Height =252
                    FontSize =9
                    TabIndex =33
                    Name ="SCHINREC"
                    ControlSource ="SCHINREC"
                    Format ="mm/dd/yy"
                    GroupTable =31

                    LayoutCachedLeft =2364
                    LayoutCachedTop =3408
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =3660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =31
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1476
                            Top =3408
                            Width =816
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Received:"
                            GroupTable =31
                            LayoutCachedLeft =1476
                            LayoutCachedTop =3408
                            LayoutCachedWidth =2292
                            LayoutCachedHeight =3660
                            LayoutGroup =5
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =31
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4092
                    Top =3408
                    Width =1176
                    Height =252
                    FontSize =9
                    TabIndex =34
                    Name ="SCHINSEXP"
                    ControlSource ="SCHINSEXP"
                    Format ="mm/dd/yy"
                    GroupTable =31

                    LayoutCachedLeft =4092
                    LayoutCachedTop =3408
                    LayoutCachedWidth =5268
                    LayoutCachedHeight =3660
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =31
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =3360
                            Top =3408
                            Width =660
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =1643706
                            Name ="Label21"
                            Caption ="Expires:"
                            GroupTable =31
                            LayoutCachedLeft =3360
                            LayoutCachedTop =3408
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =3660
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =5
                            ThemeFontIndex =1
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =31
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =300
                    Top =3525
                    Width =1050
                    Height =420
                    FontSize =8
                    FontWeight =700
                    Name ="Label22"
                    Caption ="when not set\015\012at Lease Level"
                    LayoutCachedLeft =300
                    LayoutCachedTop =3525
                    LayoutCachedWidth =1350
                    LayoutCachedHeight =3945
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9516
                    Top =1488
                    Width =4248
                    Height =840
                    FontSize =9
                    TabIndex =20
                    Name ="SchTitleHolder"
                    ControlSource ="SchTitleHolder"
                    GroupTable =23

                    LayoutCachedLeft =9516
                    LayoutCachedTop =1488
                    LayoutCachedWidth =13764
                    LayoutCachedHeight =2328
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =8436
                            Top =1488
                            Width =1008
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label34"
                            Caption ="Title Holder:"
                            GroupTable =23
                            LayoutCachedLeft =8436
                            LayoutCachedTop =1488
                            LayoutCachedWidth =9444
                            LayoutCachedHeight =1728
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1224
                    Top =96
                    Width =3576
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Schedule No. is required"
                    GroupTable =2

                    LayoutCachedLeft =1224
                    LayoutCachedTop =96
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =396
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =252
                            Top =96
                            Width =900
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label45"
                            Caption ="SchNo:"
                            GroupTable =2
                            LayoutCachedLeft =252
                            LayoutCachedTop =96
                            LayoutCachedWidth =1152
                            LayoutCachedHeight =396
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8664
                    Top =96
                    Width =3996
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SchShortDesc"
                    ControlSource ="SchShortDesc"
                    GroupTable =2

                    LayoutCachedLeft =8664
                    LayoutCachedTop =96
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =396
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =7716
                            Top =96
                            Width =876
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label48"
                            Caption ="Short Desc:"
                            GroupTable =2
                            LayoutCachedLeft =7716
                            LayoutCachedTop =96
                            LayoutCachedWidth =8592
                            LayoutCachedHeight =396
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14736
                    Top =108
                    Width =696
                    Height =228
                    FontSize =8
                    TabIndex =8
                    ForeColor =5026082
                    Name ="clientID"
                    ControlSource ="clientID"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14736
                    LayoutCachedTop =108
                    LayoutCachedWidth =15432
                    LayoutCachedHeight =336
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =14040
                            Top =108
                            Width =696
                            Height =228
                            FontSize =8
                            FontWeight =700
                            ForeColor =5026082
                            Name ="Label56"
                            Caption ="ClientID:"
                            GroupTable =13
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =14040
                            LayoutCachedTop =108
                            LayoutCachedWidth =14736
                            LayoutCachedHeight =336
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            GroupTable =13
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =8
                    ListRows =4
                    ListWidth =7200
                    Left =8196
                    Top =4056
                    Width =1608
                    Height =252
                    FontSize =9
                    TabIndex =39
                    Name ="FSLSCH"
                    ControlSource ="FSLSCH"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FSL.ID, FSL.FSLYR1, FSL.FSLYR2, FSL.FSLYR3, FSL.FSLYR4, FSL.FSLYR5, FSL.F"
                        "SLYR6, FSL.FSLYR7 FROM FSL;"
                    ColumnWidths ="360;864;864;864;864;864;864;864"
                    StatusBarText ="Link to FSL table"
                    GroupTable =35
                    AllowValueListEdits =0

                    LayoutCachedLeft =8196
                    LayoutCachedTop =4056
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =4308
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =6
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =3
                            Left =6672
                            Top =4056
                            Width =1452
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="FSLSCH_Label"
                            Caption ="FSL Sch:"
                            GroupTable =35
                            LayoutCachedLeft =6672
                            LayoutCachedTop =4056
                            LayoutCachedWidth =8124
                            LayoutCachedHeight =4308
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =8196
                    Top =3456
                    Width =5316
                    Height =252
                    FontSize =9
                    TabIndex =37
                    Name ="SchBillAdd"
                    ControlSource ="SchBillAdd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LocationType, Loc"
                        "ations.LCity, Locations.LState, Locations.LZip FROM Locations WHERE (((Locations"
                        ".LocationType)='Billing' Or (Locations.LocationType)='Garage') AND ((Locations.C"
                        "lientGroupID)=[forms]![Schedule_Form]![clientgroupid])) ORDER BY Locations.Locat"
                        "ionName, Locations.LocationType, Locations.LState;"
                    ColumnWidths ="0;2160;720;1152;360;576"
                    OnChange ="[Event Procedure]"
                    GroupTable =35
                    AllowValueListEdits =0

                    LayoutCachedLeft =8196
                    LayoutCachedTop =3456
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =3708
                    ColumnStart =1
                    ColumnEnd =7
                    LayoutGroup =6
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6672
                            Top =3456
                            Width =1452
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="Label63"
                            Caption ="Billing Address:"
                            GroupTable =35
                            LayoutCachedLeft =6672
                            LayoutCachedTop =3456
                            LayoutCachedWidth =8124
                            LayoutCachedHeight =3708
                            LayoutGroup =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =9024
                    Top =5544
                    Width =708
                    Height =288
                    TabIndex =44
                    Name ="BLCDACDT"
                    ControlSource ="BLCDACDT"
                    StatusBarText ="term starts with Accepted dt not BLCD"
                    GroupTable =35

                    LayoutCachedLeft =9024
                    LayoutCachedTop =5544
                    LayoutCachedWidth =9732
                    LayoutCachedHeight =5832
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6672
                            Top =5544
                            Width =2292
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label80"
                            Caption ="Term Starts Acceptance date."
                            GroupTable =35
                            LayoutCachedLeft =6672
                            LayoutCachedTop =5544
                            LayoutCachedWidth =8964
                            LayoutCachedHeight =5832
                            RowStart =6
                            RowEnd =6
                            ColumnEnd =1
                            LayoutGroup =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1488
                    Top =1500
                    Width =336
                    TabIndex =14
                    Name ="IntLease"
                    ControlSource ="IntLease"
                    GroupTable =19

                    LayoutCachedLeft =1488
                    LayoutCachedTop =1500
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =1740
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =408
                            Top =1500
                            Width =1020
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label87"
                            Caption ="Interim Lease"
                            GroupTable =19
                            LayoutCachedLeft =408
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1428
                            LayoutCachedHeight =1740
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11208
                    Top =5544
                    Width =2304
                    Height =288
                    FontWeight =700
                    TabIndex =45
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="SchBnkClNO"
                    ControlSource ="SchBnkClNO"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =35

                    LayoutCachedLeft =11208
                    LayoutCachedTop =5544
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =5832
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =7
                    LayoutGroup =6
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =9876
                            Top =5544
                            Width =1260
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label57"
                            Caption ="Bank Client No:"
                            GroupTable =35
                            LayoutCachedLeft =9876
                            LayoutCachedTop =5544
                            LayoutCachedWidth =11136
                            LayoutCachedHeight =5832
                            RowStart =6
                            RowEnd =6
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =2
                    Left =6672
                    Top =4368
                    Width =6840
                    Height =336
                    FontWeight =700
                    Name ="Label88"
                    Caption ="                            Customer Required Invoice \"Fields\""
                    GroupTable =35
                    LayoutCachedLeft =6672
                    LayoutCachedTop =4368
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =4704
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =7
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9876
                    Top =4776
                    Width =1260
                    Height =312
                    FontSize =9
                    TabIndex =40
                    Name ="InvFld1desc"
                    ControlSource ="InvFld1desc"
                    StatusBarText ="Lessee descriptor of number to appear on invoice"
                    GroupTable =35

                    LayoutCachedLeft =9876
                    LayoutCachedTop =4776
                    LayoutCachedWidth =11136
                    LayoutCachedHeight =5088
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11916
                    Top =4776
                    Width =1596
                    Height =312
                    FontSize =9
                    TabIndex =41
                    Name ="InvFld1"
                    ControlSource ="InvFld1"
                    StatusBarText ="number itself"
                    GroupTable =35

                    LayoutCachedLeft =11916
                    LayoutCachedTop =4776
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =5088
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =11208
                            Top =4776
                            Width =636
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label94"
                            Caption ="Value:"
                            GroupTable =35
                            LayoutCachedLeft =11208
                            LayoutCachedTop =4776
                            LayoutCachedWidth =11844
                            LayoutCachedHeight =5088
                            RowStart =4
                            RowEnd =4
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9876
                    Top =5160
                    Width =1260
                    Height =312
                    FontSize =9
                    TabIndex =42
                    Name ="InvFld2desc"
                    ControlSource ="InvFld2desc"
                    StatusBarText ="Lessee descriptor of number to appear on invoice"
                    GroupTable =35

                    LayoutCachedLeft =9876
                    LayoutCachedTop =5160
                    LayoutCachedWidth =11136
                    LayoutCachedHeight =5472
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11916
                    Top =5160
                    Width =1596
                    Height =312
                    FontSize =9
                    TabIndex =43
                    Name ="Text97"
                    ControlSource ="InvFld2"
                    StatusBarText ="number itself"
                    GroupTable =35

                    LayoutCachedLeft =11916
                    LayoutCachedTop =5160
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =5472
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =11208
                            Top =5160
                            Width =636
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label98"
                            Caption ="Value:"
                            GroupTable =35
                            LayoutCachedLeft =11208
                            LayoutCachedTop =5160
                            LayoutCachedWidth =11844
                            LayoutCachedHeight =5472
                            RowStart =5
                            RowEnd =5
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8124
                    Top =2400
                    Width =252
                    TabIndex =23
                    Name ="SchTitleToTrust"
                    ControlSource ="SchTitleToTrust"
                    GroupTable =23

                    LayoutCachedLeft =8124
                    LayoutCachedTop =2400
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =2640
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6696
                            Top =2400
                            Width =1368
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label101"
                            Caption ="FA Title Trust:"
                            GroupTable =23
                            LayoutCachedLeft =6696
                            LayoutCachedTop =2400
                            LayoutCachedWidth =8064
                            LayoutCachedHeight =2640
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =4
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4092
                    Top =3732
                    Width =1176
                    Height =312
                    FontSize =9
                    TabIndex =35
                    Name ="SchSLVRec"
                    ControlSource ="SchSLVRec"
                    Format ="mm/dd/yy"
                    StatusBarText ="SLV Tables received"
                    GroupTable =31

                    LayoutCachedLeft =4092
                    LayoutCachedTop =3732
                    LayoutCachedWidth =5268
                    LayoutCachedHeight =4044
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =31
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1476
                            Top =3732
                            Width =2544
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label109"
                            Caption ="Stip Loss Table Received:"
                            GroupTable =31
                            LayoutCachedLeft =1476
                            LayoutCachedTop =3732
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =4044
                            RowStart =1
                            RowEnd =1
                            ColumnEnd =2
                            LayoutGroup =5
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =31
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5244
                    Top =1500
                    Width =1068
                    FontSize =9
                    TabIndex =15
                    Name ="SchedDt"
                    ControlSource ="SchedDt"
                    Format ="Short Date"
                    GroupTable =19

                    LayoutCachedLeft =5244
                    LayoutCachedTop =1500
                    LayoutCachedWidth =6312
                    LayoutCachedHeight =1740
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3912
                            Top =1500
                            Width =1260
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label115"
                            Caption ="Sch Date"
                            GroupTable =19
                            LayoutCachedLeft =3912
                            LayoutCachedTop =1500
                            LayoutCachedWidth =5172
                            LayoutCachedHeight =1740
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =3
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =163
                    Top =6240
                    Width =17790
                    Height =8205
                    TabIndex =55
                    Name ="TabCtl116"

                    LayoutCachedTop =6240
                    LayoutCachedWidth =17790
                    LayoutCachedHeight =14445
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =6720
                            Width =17520
                            Height =7590
                            Name ="Page117"
                            Caption ="ATF Revisions"
                            LayoutCachedLeft =135
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =135
                                    Top =7065
                                    Width =17520
                                    Height =7245
                                    Name ="ATFRevF"
                                    SourceObject ="Form.ATFRevF"
                                    LinkChildFields ="SchID"
                                    LinkMasterFields ="SchID"

                                    LayoutCachedLeft =135
                                    LayoutCachedTop =7065
                                    LayoutCachedWidth =17655
                                    LayoutCachedHeight =14310
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =6720
                            Width =17520
                            Height =7590
                            Name ="CoLessee1"
                            Caption ="Co/Sub Lessee"
                            LayoutCachedLeft =135
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =210
                                    Top =6990
                                    Width =9615
                                    Height =4380
                                    Name ="Schedule_CoLessee"
                                    SourceObject ="Form.Schedule_CoLessee"
                                    LinkChildFields ="ScheduleID"
                                    LinkMasterFields ="SchID"

                                    LayoutCachedLeft =210
                                    LayoutCachedTop =6990
                                    LayoutCachedWidth =9825
                                    LayoutCachedHeight =11370
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =130
                            Top =6720
                            Width =17525
                            Height =7590
                            Name ="Page118"
                            Caption ="Schedule Amendments"
                            LayoutCachedLeft =130
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =130
                                    Top =7064
                                    Width =16065
                                    Height =5535
                                    Name ="SchAmendf"
                                    SourceObject ="Form.ScheduleAmends"
                                    LinkChildFields ="ScheduleID"
                                    LinkMasterFields ="SchID"

                                    LayoutCachedLeft =130
                                    LayoutCachedTop =7064
                                    LayoutCachedWidth =16195
                                    LayoutCachedHeight =12599
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =6720
                            Width =17520
                            Height =7590
                            Name ="Permanent Funding Info"
                            EventProcPrefix ="Permanent_Funding_Info"
                            LayoutCachedLeft =135
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =435
                                    Top =7035
                                    Width =2745
                                    Height =405
                                    FontSize =16
                                    FontWeight =700
                                    Name ="Label836"
                                    Caption ="Permanent Funding"
                                    LayoutCachedLeft =435
                                    LayoutCachedTop =7035
                                    LayoutCachedWidth =3180
                                    LayoutCachedHeight =7440
                                    ForeThemeColorIndex =6
                                    ForeShade =75.0
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6288
                                    Top =7476
                                    Height =276
                                    FontSize =9
                                    TabIndex =1
                                    Name ="SchFEcoRec"
                                    ControlSource ="SchFEcoRec"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="date final economics received"
                                    GroupTable =55

                                    LayoutCachedLeft =6288
                                    LayoutCachedTop =7476
                                    LayoutCachedWidth =7728
                                    LayoutCachedHeight =7752
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =7
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =55
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =4128
                                            Top =7476
                                            Width =2088
                                            Height =276
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label145"
                                            Caption ="Synd Economics Received:"
                                            GroupTable =55
                                            LayoutCachedLeft =4128
                                            LayoutCachedTop =7476
                                            LayoutCachedWidth =6216
                                            LayoutCachedHeight =7752
                                            LayoutGroup =7
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =6
                                            ForeShade =50.0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =55
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6288
                                    Top =7824
                                    Height =276
                                    FontSize =9
                                    TabIndex =2
                                    Name ="SchFEcoA"
                                    ControlSource ="SchFEcoA"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="date final economics apporved"
                                    GroupTable =55

                                    LayoutCachedLeft =6288
                                    LayoutCachedTop =7824
                                    LayoutCachedWidth =7728
                                    LayoutCachedHeight =8100
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =7
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =55
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =4128
                                            Top =7824
                                            Width =2088
                                            Height =276
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label146"
                                            Caption ="Synd Economics Approved:"
                                            GroupTable =55
                                            LayoutCachedLeft =4128
                                            LayoutCachedTop =7824
                                            LayoutCachedWidth =6216
                                            LayoutCachedHeight =8100
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =7
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =6
                                            ForeShade =50.0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =55
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =6288
                                    Top =8172
                                    Height =276
                                    FontSize =9
                                    TabIndex =3
                                    Name ="Text151"
                                    ControlSource ="SchFPkgO"
                                    StatusBarText ="date Funding package sent to bank"
                                    GroupTable =55

                                    LayoutCachedLeft =6288
                                    LayoutCachedTop =8172
                                    LayoutCachedWidth =7728
                                    LayoutCachedHeight =8448
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =7
                                    BorderThemeColorIndex =6
                                    BorderShade =50.0
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =55
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =4128
                                            Top =8172
                                            Width =2088
                                            Height =276
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label152"
                                            Caption ="Perm Funding Pkg. Out:"
                                            GroupTable =55
                                            LayoutCachedLeft =4128
                                            LayoutCachedTop =8172
                                            LayoutCachedWidth =6216
                                            LayoutCachedHeight =8448
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =7
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =6
                                            ForeShade =50.0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =55
                                        End
                                    End
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =420
                                    Top =7740
                                    Width =3024
                                    Height =1920
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =4
                                    ForeColor =7500402
                                    BorderColor =15921906
                                    Name ="GroupFundList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT SchGrp.SGrpNO AS [Group], SchGrp.Assignee_Bank AS Assignee, SchGrp.Assign"
                                        "mentDate AS [Fund Date], SchGrp.Assigned_Indicator AS [P/F], SchGrp.SGrpID, SchG"
                                        "rp.SGrpSchID FROM SchGrp WHERE (((SchGrp.SGrpSchID)=[Forms]![Schedule_Form]![Sch"
                                        "ID]));"
                                    ColumnWidths ="720;864;1008;288;0;0"
                                    OnDblClick ="[Event Procedure]"

                                    LayoutCachedLeft =420
                                    LayoutCachedTop =7740
                                    LayoutCachedWidth =3444
                                    LayoutCachedHeight =9660
                                    BackThemeColorIndex =-1
                                    BorderShade =95.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =420
                                            Top =7500
                                            Width =2760
                                            Height =315
                                            FontWeight =700
                                            Name ="Label961"
                                            Caption ="Group Funded Date P/F"
                                            LayoutCachedLeft =420
                                            LayoutCachedTop =7500
                                            LayoutCachedWidth =3180
                                            LayoutCachedHeight =7815
                                            ForeThemeColorIndex =6
                                            ForeShade =75.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13224
                                    Top =7560
                                    Width =1596
                                    Height =255
                                    FontSize =9
                                    TabIndex =5
                                    Name ="SchCurrRent"
                                    ControlSource ="SchCurrRent"
                                    Format ="$#,##0.00;($#,##0.00)"
                                    StatusBarText ="current rent for sliver calculation"
                                    ValidationRule =">=0"
                                    ValidationText ="Per Unit Curr Rent  must be >=0"
                                    DefaultValue ="0"
                                    GroupTable =63

                                    LayoutCachedLeft =13224
                                    LayoutCachedTop =7560
                                    LayoutCachedWidth =14820
                                    LayoutCachedHeight =7815
                                    ColumnStart =6
                                    ColumnEnd =7
                                    LayoutGroup =11
                                    BorderThemeColorIndex =2
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =63
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =8652
                                            Top =7560
                                            Width =4500
                                            Height =255
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label81"
                                            Caption ="Per Unit CurrRent:(Sliver Rent Calculation):"
                                            GroupTable =63
                                            LayoutCachedLeft =8652
                                            LayoutCachedTop =7560
                                            LayoutCachedWidth =13152
                                            LayoutCachedHeight =7815
                                            ColumnEnd =5
                                            LayoutGroup =11
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =63
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13224
                                    Top =7884
                                    Width =1596
                                    Height =252
                                    FontSize =9
                                    TabIndex =6
                                    Name ="SchCRExp"
                                    ControlSource ="SchCRExp"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="current rent expiry"
                                    GroupTable =63

                                    LayoutCachedLeft =13224
                                    LayoutCachedTop =7884
                                    LayoutCachedWidth =14820
                                    LayoutCachedHeight =8136
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =6
                                    ColumnEnd =7
                                    LayoutGroup =11
                                    BorderThemeColorIndex =2
                                    ThemeFontIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =63
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =11196
                                            Top =7884
                                            Width =1956
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label82"
                                            Caption ="Old Lease Expiry:"
                                            GroupTable =63
                                            LayoutCachedLeft =11196
                                            LayoutCachedTop =7884
                                            LayoutCachedWidth =13152
                                            LayoutCachedHeight =8136
                                            RowStart =1
                                            RowEnd =1
                                            ColumnStart =4
                                            ColumnEnd =5
                                            LayoutGroup =11
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =63
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =8652
                                    Top =7884
                                    Width =2460
                                    Height =252
                                    Name ="EmptyCell1196"
                                    GroupTable =63
                                    RightPadding =0
                                    LayoutCachedLeft =8652
                                    LayoutCachedTop =7884
                                    LayoutCachedWidth =11112
                                    LayoutCachedHeight =8136
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11112
                                    Top =7884
                                    Width =0
                                    Height =252
                                    Name ="EmptyCell1197"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11112
                                    LayoutCachedTop =7884
                                    LayoutCachedWidth =11112
                                    LayoutCachedHeight =8136
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11112
                                    Top =7884
                                    Width =48
                                    Height =252
                                    Name ="EmptyCell1198"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11112
                                    LayoutCachedTop =7884
                                    LayoutCachedWidth =11160
                                    LayoutCachedHeight =8136
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11160
                                    Top =7884
                                    Width =0
                                    Height =252
                                    Name ="EmptyCell1199"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11160
                                    LayoutCachedTop =7884
                                    LayoutCachedWidth =11160
                                    LayoutCachedHeight =8136
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =135
                            Top =6720
                            Width =17520
                            Height =7590
                            Name ="BillingRules"
                            Caption ="Billing Rules"
                            LayoutCachedLeft =135
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =195
                                    Top =7050
                                    Width =16605
                                    Height =6090
                                    Name ="ScheduleBillingRules_subform"
                                    SourceObject ="Form.ScheduleBillingRules_subform"
                                    LinkChildFields ="ScheduleId"
                                    LinkMasterFields ="SchID"

                                    LayoutCachedLeft =195
                                    LayoutCachedTop =7050
                                    LayoutCachedWidth =16800
                                    LayoutCachedHeight =13140
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =6720
                            Width =17535
                            Height =7590
                            Name ="Documentation"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6720
                            LayoutCachedWidth =17655
                            LayoutCachedHeight =14310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =120
                                    Top =7260
                                    Width =13740
                                    Height =6900
                                    Name ="DocLinksSub"
                                    SourceObject ="Form.DocLinksSub"

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =7260
                                    LayoutCachedWidth =13860
                                    LayoutCachedHeight =14160
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =120
                                    Top =6840
                                    Width =3300
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="AddDocLink"
                                    Caption ="Add Another Document Link"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =6840
                                    LayoutCachedWidth =3420
                                    LayoutCachedHeight =7140
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
                            End
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3216
                    Top =468
                    Width =576
                    Height =252
                    TabIndex =4
                    Name ="SchAH"
                    ControlSource ="SchAH"
                    StatusBarText ="Autohauler schedule"
                    GroupTable =2

                    LayoutCachedLeft =3216
                    LayoutCachedTop =468
                    LayoutCachedWidth =3792
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1896
                            Top =468
                            Width =1260
                            Height =252
                            FontSize =8
                            FontWeight =700
                            Name ="Label121"
                            Caption ="Autohaulers:"
                            GroupTable =2
                            LayoutCachedLeft =1896
                            LayoutCachedTop =468
                            LayoutCachedWidth =3156
                            LayoutCachedHeight =720
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =13236
                    Top =3768
                    Width =276
                    Height =216
                    TabIndex =38
                    Name ="SchFABill"
                    ControlSource ="SchFABill"
                    StatusBarText ="FA Bills an assigned schedule on behalf of bank"
                    GroupTable =35

                    LayoutCachedLeft =13236
                    LayoutCachedTop =3768
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =3984
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =11208
                            Top =3768
                            Width =1968
                            Height =216
                            FontSize =8
                            FontWeight =700
                            Name ="Label125"
                            Caption ="FA Bills after Assignment:"
                            GroupTable =35
                            LayoutCachedLeft =11208
                            LayoutCachedTop =3768
                            LayoutCachedWidth =13176
                            LayoutCachedHeight =3984
                            RowStart =1
                            RowEnd =1
                            ColumnStart =5
                            ColumnEnd =6
                            LayoutGroup =6
                            ThemeFontIndex =1
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =35
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1224
                    Top =468
                    Width =612
                    Height =252
                    TabIndex =3
                    Name ="SchExSyn"
                    ControlSource ="SchExSyn"
                    StatusBarText ="schedule is not to be syndicated added 5-17-13"
                    GroupTable =2

                    LayoutCachedLeft =1224
                    LayoutCachedTop =468
                    LayoutCachedWidth =1836
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =252
                            Top =468
                            Width =900
                            Height =252
                            FontSize =8
                            FontWeight =700
                            Name ="Label126"
                            Caption ="Do Not Syn: "
                            GroupTable =2
                            LayoutCachedLeft =252
                            LayoutCachedTop =468
                            LayoutCachedWidth =1152
                            LayoutCachedHeight =720
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =14310
                    Top =1245
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =18
                    Name ="Command140"
                    Caption ="             Print RAL Attachment"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="AssignNtAtt1"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command140\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argum"
                                "ent Name=\"ReportName\">AssignNtA"
                        End
                        Begin
                            Comment ="_AXL:tt1</Argument><Argument Name=\"View\">Print Preview</Argument></Action></St"
                                "atements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14310
                    LayoutCachedTop =1245
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =1530
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =5652
                    Top =96
                    Width =1992
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SchFAID"
                    ControlSource ="SchFAID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Deals.faid, Deals.dealShortDesc From Deals WHERE ClientRef=3410 ORDER BY "
                        "Deals.faid DESC;"
                    ColumnWidths ="720;4320"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="FAID is required"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =5652
                    LayoutCachedTop =96
                    LayoutCachedWidth =7644
                    LayoutCachedHeight =396
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =4872
                            Top =96
                            Width =708
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label110"
                            Caption ="FAID:"
                            GroupTable =2
                            LayoutCachedLeft =4872
                            LayoutCachedTop =96
                            LayoutCachedWidth =5580
                            LayoutCachedHeight =396
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =14310
                    Top =1590
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =26
                    Name ="RALExport"
                    Caption ="  RAL Data (Export)"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
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

                    LayoutCachedLeft =14310
                    LayoutCachedTop =1590
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =1875
                    PictureCaptionArrangement =5
                    RowStart =1
                    RowEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =14310
                    Top =3330
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =32
                    Name ="MUVPmtBtn"
                    Caption ="Main Unit Invoice Update"
                    ControlTipText ="Brings up main unit plus child parts from same vendor"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="MainUnitPmtF"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"MUVPmtBtn\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">MainUnitPmtF</"
                        End
                        Begin
                            Comment ="_AXL:Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14310
                    LayoutCachedTop =3330
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =3615
                    RowStart =5
                    RowEnd =5
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =14310
                    Top =3675
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =46
                    Name ="cvinvupdtbtn"
                    Caption ="Child Vendor Invoice Update"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Child Vendors Only"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cvinvupdtbtn\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argum"
                                "ent Name=\"FormName\">Child Vendo"
                        End
                        Begin
                            Comment ="_AXL:rs Only</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14310
                    LayoutCachedTop =3675
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =3960
                    RowStart =6
                    RowEnd =6
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =14310
                    Top =4020
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =47
                    Name ="AddFHUT"
                    Caption ="Add/Update FHUT Records"
                    StatusBarText ="Each time this button is clicked a FHUT record will be added to each unit on the"
                        " schedule"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Each time this button is clicked a FHUT record will be added to each unit on the"
                        " schedule"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =14310
                    LayoutCachedTop =4020
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =4305
                    RowStart =7
                    RowEnd =7
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =14175
                    Top =1005
                    Width =3495
                    Height =2010
                    Name ="Box158"
                    LayoutCachedLeft =14175
                    LayoutCachedTop =1005
                    LayoutCachedWidth =17670
                    LayoutCachedHeight =3015
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =14280
                    Top =855
                    Width =690
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label162"
                    Caption ="DOCS"
                    LayoutCachedLeft =14280
                    LayoutCachedTop =855
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =1080
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13344
                    Top =108
                    Width =696
                    Height =228
                    FontSize =8
                    TabIndex =7
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="SchMLRef"
                    StatusBarText ="reference to master lease"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13344
                    LayoutCachedTop =108
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =336
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12648
                            Top =108
                            Width =696
                            Height =228
                            FontSize =8
                            FontWeight =700
                            ForeColor =5026082
                            Name ="Label129"
                            Caption ="MLID"
                            GroupTable =13
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =12648
                            LayoutCachedTop =108
                            LayoutCachedWidth =13344
                            LayoutCachedHeight =336
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =13
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =15975
                    Top =4860
                    Width =1620
                    Height =345
                    FontSize =10
                    TabIndex =50
                    Name ="UPDtFHUT"
                    Caption ="Update FHUT"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =15975
                    LayoutCachedTop =4860
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =5205
                    RowStart =8
                    RowEnd =8
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =14310
                    Top =1935
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =27
                    Name ="Command164"
                    Caption ="      Print Title Status"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="SchVinReport"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command164\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argum"
                                "ent Name=\"ReportName\">SchVinRep"
                        End
                        Begin
                            Comment ="_AXL:ort</Argument><Argument Name=\"View\">Print Preview</Argument></Action></St"
                                "atements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14310
                    LayoutCachedTop =1935
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =2220
                    RowStart =2
                    RowEnd =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16128
                    Top =108
                    Width =696
                    Height =228
                    FontSize =8
                    TabIndex =9
                    ForeColor =5026082
                    Name ="SchID"
                    ControlSource ="SchID"
                    GroupTable =13
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16128
                    LayoutCachedTop =108
                    LayoutCachedWidth =16824
                    LayoutCachedHeight =336
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =15432
                            Top =108
                            Width =696
                            Height =228
                            FontSize =8
                            FontWeight =700
                            ForeColor =5026082
                            Name ="Label130"
                            Caption ="SchID"
                            GroupTable =13
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =15432
                            LayoutCachedTop =108
                            LayoutCachedWidth =16128
                            LayoutCachedHeight =336
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            GroupTable =13
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =4452
                    Top =468
                    Width =348
                    Height =252
                    TabIndex =5
                    BorderColor =8210719
                    Name ="SchFSL"
                    ControlSource ="SchFSL"
                    StatusBarText ="FSL Lease"
                    GroupTable =2

                    LayoutCachedLeft =4452
                    LayoutCachedTop =468
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3852
                            Top =468
                            Width =540
                            Height =252
                            FontSize =8
                            FontWeight =700
                            Name ="Label46"
                            Caption ="FSL:"
                            GroupTable =2
                            LayoutCachedLeft =3852
                            LayoutCachedTop =468
                            LayoutCachedWidth =4392
                            LayoutCachedHeight =720
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =420
                    Top =2724
                    Width =690
                    Name ="EmptyCell284"
                    GroupTable =16
                    LayoutCachedLeft =420
                    LayoutCachedTop =2724
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =420
                    Top =2412
                    Width =690
                    Name ="EmptyCell287"
                    GroupTable =16
                    LayoutCachedLeft =420
                    LayoutCachedTop =2412
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =2652
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =255
                    Top =1005
                    Width =6105
                    Height =2010
                    Name ="Box332"
                    LayoutCachedLeft =255
                    LayoutCachedTop =1005
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =3015
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextFontCharSet =163
                    TextAlign =2
                    Left =390
                    Top =840
                    Width =2475
                    Height =285
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Document Info"
                    LayoutCachedLeft =390
                    LayoutCachedTop =840
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =1125
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =6540
                    Top =1005
                    Width =7410
                    Height =2010
                    Name ="Box333"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =1005
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =3015
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =6668
                    Top =840
                    Width =2475
                    Height =285
                    FontWeight =700
                    Name ="Label334"
                    Caption ="Equipment Title Info:"
                    LayoutCachedLeft =6668
                    LayoutCachedTop =840
                    LayoutCachedWidth =9143
                    LayoutCachedHeight =1125
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin EmptyCell
                    Left =8436
                    Top =1788
                    Width =1008
                    Name ="EmptyCell392"
                    GroupTable =23
                    LayoutCachedLeft =8436
                    LayoutCachedTop =1788
                    LayoutCachedWidth =9444
                    LayoutCachedHeight =2028
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =8436
                    Top =2088
                    Width =1008
                    Name ="EmptyCell393"
                    GroupTable =23
                    LayoutCachedLeft =8436
                    LayoutCachedTop =2088
                    LayoutCachedWidth =9444
                    LayoutCachedHeight =2328
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =8436
                    Top =2700
                    Width =1008
                    Name ="EmptyCell401"
                    GroupTable =23
                    LayoutCachedLeft =8436
                    LayoutCachedTop =2700
                    LayoutCachedWidth =9444
                    LayoutCachedHeight =2940
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =9516
                    Top =2700
                    Width =4248
                    Name ="EmptyCell402"
                    GroupTable =23
                    LayoutCachedLeft =9516
                    LayoutCachedTop =2700
                    LayoutCachedWidth =13764
                    LayoutCachedHeight =2940
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =240
                    Top =3240
                    Width =6105
                    Height =1215
                    Name ="Box403"
                    LayoutCachedLeft =240
                    LayoutCachedTop =3240
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =4455
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =360
                    Top =3075
                    Width =2475
                    Height =285
                    FontWeight =700
                    Name ="Label404"
                    Caption ="Insurance Info"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3075
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =3360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin EmptyCell
                    Left =1476
                    Top =4116
                    Width =816
                    Height =300
                    Name ="EmptyCell437"
                    GroupTable =31
                    LayoutCachedLeft =1476
                    LayoutCachedTop =4116
                    LayoutCachedWidth =2292
                    LayoutCachedHeight =4416
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =5
                    GroupTable =31
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =6525
                    Top =3225
                    Width =7425
                    Height =2595
                    Name ="Box447"
                    LayoutCachedLeft =6525
                    LayoutCachedTop =3225
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =5820
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =6630
                    Top =3060
                    Width =2475
                    Height =285
                    FontWeight =700
                    Name ="Label448"
                    Caption ="'Billing' Info"
                    LayoutCachedLeft =6630
                    LayoutCachedTop =3060
                    LayoutCachedWidth =9105
                    LayoutCachedHeight =3345
                    BackThemeColorIndex =6
                    BackShade =75.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                End
                Begin EmptyCell
                    Left =6672
                    Top =3768
                    Width =4464
                    Height =216
                    Name ="EmptyCell561"
                    GroupTable =35
                    LayoutCachedLeft =6672
                    LayoutCachedTop =3768
                    LayoutCachedWidth =11136
                    LayoutCachedHeight =3984
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =4
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =6672
                    Top =4776
                    Width =3060
                    Height =312
                    FontSize =9
                    FontWeight =700
                    Name ="Label89"
                    Caption ="Description/Label 1:"
                    GroupTable =35
                    LayoutCachedLeft =6672
                    LayoutCachedTop =4776
                    LayoutCachedWidth =9732
                    LayoutCachedHeight =5088
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =6672
                    Top =5160
                    Width =3060
                    Height =312
                    FontSize =9
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Description/Label 2:"
                    GroupTable =35
                    LayoutCachedLeft =6672
                    LayoutCachedTop =5160
                    LayoutCachedWidth =9732
                    LayoutCachedHeight =5472
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    LayoutGroup =6
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =35
                End
                Begin EmptyCell
                    Left =9792
                    Top =4776
                    Width =12
                    Height =312
                    Name ="EmptyCell604"
                    GroupTable =35
                    LayoutCachedLeft =9792
                    LayoutCachedTop =4776
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =5088
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin EmptyCell
                    Left =9792
                    Top =5160
                    Width =12
                    Height =312
                    Name ="EmptyCell605"
                    GroupTable =35
                    LayoutCachedLeft =9792
                    LayoutCachedTop =5160
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =5472
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =240
                    Top =4710
                    Width =6105
                    Height =1095
                    Name ="Box643"
                    LayoutCachedLeft =240
                    LayoutCachedTop =4710
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =5805
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =330
                    Top =4545
                    Width =2505
                    Height =285
                    FontWeight =700
                    Name ="Label644"
                    Caption ="RAL Info"
                    LayoutCachedLeft =330
                    LayoutCachedTop =4545
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =4830
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin EmptyCell
                    Left =9792
                    Top =5544
                    Width =12
                    Height =288
                    Name ="EmptyCell675"
                    GroupTable =35
                    LayoutCachedLeft =9792
                    LayoutCachedTop =5544
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =5832
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =14175
                    Top =3225
                    Width =3495
                    Height =3000
                    Name ="Box679"
                    LayoutCachedLeft =14175
                    LayoutCachedTop =3225
                    LayoutCachedWidth =17670
                    LayoutCachedHeight =6225
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =14340
                    Top =3075
                    Width =705
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label680"
                    Caption ="OPS"
                    LayoutCachedLeft =14340
                    LayoutCachedTop =3075
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =3300
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14310
                    Top =4440
                    Width =1365
                    Height =255
                    FontSize =10
                    TabIndex =48
                    Name ="RentReport"
                    Caption ="Rent Report ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14310
                    LayoutCachedTop =4440
                    LayoutCachedWidth =15675
                    LayoutCachedHeight =4695
                    RowStart =9
                    RowEnd =9
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =15765
                    Top =4470
                    Width =1620
                    Height =255
                    FontSize =8
                    TabIndex =49
                    Name ="Location"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Count(vw_SixKeys.[UnitID]) "
                        "AS Units, Locations.LocationType, vw_SixKeys.mlorig FROM Locations INNER JOIN vw"
                        "_SixKeys ON Locations.LocationID = vw_SixKeys.LocationID GROUP BY Locations.Loca"
                        "tionID, Locations.LocationName, Locations.LocationType, Locations.LocObsolete, L"
                        "ocations.ClientREF, vw_SixKeys.mlorig HAVING (((Locations.LocObsolete)=No) AND ("
                        "(Locations.ClientREF)=[Forms]![Schedule_Form]![clientID]) AND ((vw_SixKeys.mlori"
                        "g)=\"FA\")) ORDER BY Locations.LocationName, Locations.LocationType DESC;"
                    ColumnWidths ="0;1872;720"
                    StatusBarText ="Link to Contacts table"
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =15765
                    LayoutCachedTop =4470
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =4725
                    RowStart =10
                    RowEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16104
                    Top =336
                    Width =696
                    Height =228
                    FontSize =8
                    TabIndex =11
                    ForeColor =5026082
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    GroupTable =57
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16104
                    LayoutCachedTop =336
                    LayoutCachedWidth =16800
                    LayoutCachedHeight =564
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =8
                    GroupTable =57
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =15408
                            Top =336
                            Width =696
                            Height =228
                            FontSize =8
                            FontWeight =700
                            ForeColor =5026082
                            Name ="Label844"
                            Caption ="MLOrig"
                            GroupTable =57
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =15408
                            LayoutCachedTop =336
                            LayoutCachedWidth =16104
                            LayoutCachedHeight =564
                            LayoutGroup =8
                            GroupTable =57
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14310
                    Top =5565
                    Width =1575
                    Height =285
                    FontWeight =700
                    TabIndex =53
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="FAVendorNumber"
                    ControlSource ="FAVendorNumber"

                    LayoutCachedLeft =14310
                    LayoutCachedTop =5565
                    LayoutCachedWidth =15885
                    LayoutCachedHeight =5850
                    ColumnStart =1
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =14310
                            Top =5340
                            Width =1575
                            Height =225
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label846"
                            Caption ="*FA Vendor #"
                            LayoutCachedLeft =14310
                            LayoutCachedTop =5340
                            LayoutCachedWidth =15885
                            LayoutCachedHeight =5565
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15975
                    Top =5565
                    Width =1575
                    Height =285
                    FontWeight =700
                    TabIndex =54
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="FAPOnumber"
                    ControlSource ="FAPOnumber"

                    LayoutCachedLeft =15975
                    LayoutCachedTop =5565
                    LayoutCachedWidth =17550
                    LayoutCachedHeight =5850
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =14310
                    Top =5895
                    Width =3120
                    Height =210
                    FontSize =8
                    FontWeight =700
                    Name ="Label853"
                    Caption ="*Assigned to FA by Client (billing)"
                    LayoutCachedLeft =14310
                    LayoutCachedTop =5895
                    LayoutCachedWidth =17430
                    LayoutCachedHeight =6105
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2784
                    Top =1176
                    Width =1068
                    Height =252
                    FontSize =9
                    TabIndex =12
                    Name ="ApprovalRequest"
                    ControlSource ="ApprovalRequest"
                    Format ="Short Date"
                    StatusBarText ="Internal Approval Request Date"
                    GroupTable =19

                    LayoutCachedLeft =2784
                    LayoutCachedTop =1176
                    LayoutCachedWidth =3852
                    LayoutCachedHeight =1428
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BackThemeColorIndex =7
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextAlign =3
                            Left =408
                            Top =1176
                            Width =2304
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="Label855"
                            Caption ="Schedule Draft Out "
                            GroupTable =19
                            LayoutCachedLeft =408
                            LayoutCachedTop =1176
                            LayoutCachedWidth =2712
                            LayoutCachedHeight =1428
                            ColumnEnd =2
                            LayoutGroup =3
                            ThemeFontIndex =1
                            BackThemeColorIndex =7
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5244
                    Top =1176
                    Width =1068
                    Height =252
                    FontSize =9
                    TabIndex =13
                    Name ="ApprovalReply"
                    ControlSource ="ApprovalReply"
                    Format ="Short Date"
                    StatusBarText ="Internal Approval REPLY Date"
                    GroupTable =19

                    LayoutCachedLeft =5244
                    LayoutCachedTop =1176
                    LayoutCachedWidth =6312
                    LayoutCachedHeight =1428
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    BackThemeColorIndex =7
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =19
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextAlign =3
                            Left =3912
                            Top =1176
                            Width =1260
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="Label857"
                            Caption ="Draft Approved "
                            GroupTable =19
                            LayoutCachedLeft =3912
                            LayoutCachedTop =1176
                            LayoutCachedWidth =5172
                            LayoutCachedHeight =1428
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =3
                            ThemeFontIndex =1
                            BackThemeColorIndex =7
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =19
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14310
                    Top =2280
                    Width =3060
                    Height =285
                    FontSize =10
                    TabIndex =31
                    Name ="UnitSummary"
                    Caption =" Sch Unit Summary"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
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

                    LayoutCachedLeft =14310
                    LayoutCachedTop =2280
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =2565
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1188
                    Top =2100
                    Width =5088
                    FontSize =8
                    TabIndex =28
                    Name ="SchSigner"
                    ControlSource ="SchSignerID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT contactID, FullName, Title, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3410 Order by FullName;"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =16

                    LayoutCachedLeft =1188
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6276
                    LayoutCachedHeight =2340
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =16
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =420
                            Top =2100
                            Width =690
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label1001"
                            Caption ="Signers"
                            GroupTable =16
                            LayoutCachedLeft =420
                            LayoutCachedTop =2100
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =2340
                            LayoutGroup =9
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =16
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1188
                    Top =2412
                    Width =5088
                    FontSize =8
                    TabIndex =29
                    Name ="SchSigner2"
                    ControlSource ="SchSigner2ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT contactID, FullName, Title, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3410 Order by FullName;"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =16

                    LayoutCachedLeft =1188
                    LayoutCachedTop =2412
                    LayoutCachedWidth =6276
                    LayoutCachedHeight =2652
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =16
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =1188
                    Top =2724
                    Width =5088
                    FontSize =8
                    TabIndex =30
                    Name ="SchSigner3"
                    ControlSource ="SchSigner3ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT contactID, FullName, Title, ClientID FROM ClientContacts_Base WHERE Clien"
                        "tID= 3410 Order by FullName;"
                    ColumnWidths ="0;2880;1440"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =16

                    LayoutCachedLeft =1188
                    LayoutCachedTop =2724
                    LayoutCachedWidth =6276
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =16
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =15975
                    Top =5340
                    Width =1575
                    Height =225
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Label851"
                    Caption ="*FA PO Nunber"
                    LayoutCachedLeft =15975
                    LayoutCachedTop =5340
                    LayoutCachedWidth =17550
                    LayoutCachedHeight =5565
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =7716
                    Top =468
                    Width =876
                    Height =252
                    Name ="EmptyCell1088"
                    GroupTable =2
                    LayoutCachedLeft =7716
                    LayoutCachedTop =468
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8664
                    Top =468
                    Width =3996
                    Height =252
                    Name ="EmptyCell1089"
                    GroupTable =2
                    LayoutCachedLeft =8664
                    LayoutCachedTop =468
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6336
                    Top =2412
                    Width =96
                    Name ="EmptyCell1098"
                    GroupTable =16
                    LayoutCachedLeft =6336
                    LayoutCachedTop =2412
                    LayoutCachedWidth =6432
                    LayoutCachedHeight =2652
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =6336
                    Top =2724
                    Width =96
                    Name ="EmptyCell1099"
                    GroupTable =16
                    LayoutCachedLeft =6336
                    LayoutCachedTop =2724
                    LayoutCachedWidth =6432
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =6336
                    Top =2100
                    Width =96
                    Name ="EmptyCell1100"
                    GroupTable =16
                    LayoutCachedLeft =6336
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6432
                    LayoutCachedHeight =2340
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1159
                    Top =4935
                    Width =1350
                    Height =780
                    FontSize =9
                    TabIndex =51
                    Name ="Text1101"
                    ControlSource ="RALOut"
                    Format ="Short Date"
                    StatusBarText ="date schedule sent out"

                    LayoutCachedLeft =1159
                    LayoutCachedTop =4935
                    LayoutCachedWidth =2509
                    LayoutCachedHeight =5715
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =435
                            Top =4935
                            Width =724
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label1102"
                            Caption ="RAL Out"
                            LayoutCachedLeft =435
                            LayoutCachedTop =4935
                            LayoutCachedWidth =1159
                            LayoutCachedHeight =5205
                            ColumnStart =1
                            ColumnEnd =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3593
                    Top =4935
                    Width =1350
                    Height =780
                    FontSize =9
                    TabIndex =52
                    Name ="Text1103"
                    ControlSource ="RALBack"
                    Format ="Short Date"
                    StatusBarText ="date schedule in"

                    LayoutCachedLeft =3593
                    LayoutCachedTop =4935
                    LayoutCachedWidth =4943
                    LayoutCachedHeight =5715
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =2824
                            Top =4935
                            Width =769
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label1104"
                            Caption ="RAL Back"
                            LayoutCachedLeft =2824
                            LayoutCachedTop =4935
                            LayoutCachedWidth =3593
                            LayoutCachedHeight =5205
                            ColumnStart =3
                            ColumnEnd =3
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin EmptyCell
                    Left =1488
                    Top =1812
                    Width =336
                    Height =276
                    Name ="EmptyCell1182"
                    GroupTable =19
                    LayoutCachedLeft =1488
                    LayoutCachedTop =1812
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =2088
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =408
                    Top =1812
                    Width =1020
                    Height =276
                    Name ="EmptyCell1183"
                    GroupTable =19
                    LayoutCachedLeft =408
                    LayoutCachedTop =1812
                    LayoutCachedWidth =1428
                    LayoutCachedHeight =2088
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =2784
                    Top =1500
                    Width =1068
                    Name ="EmptyCell1190"
                    GroupTable =19
                    LayoutCachedLeft =2784
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3852
                    LayoutCachedHeight =1740
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =1884
                    Top =1500
                    Width =828
                    Name ="EmptyCell1191"
                    GroupTable =19
                    LayoutCachedLeft =1884
                    LayoutCachedTop =1500
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =1740
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =9876
                    Top =4056
                    Width =1968
                    Height =252
                    Name ="EmptyCell1194"
                    GroupTable =35
                    LayoutCachedLeft =9876
                    LayoutCachedTop =4056
                    LayoutCachedWidth =11844
                    LayoutCachedHeight =4308
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin EmptyCell
                    Left =11916
                    Top =4056
                    Width =1596
                    Height =252
                    Name ="EmptyCell1195"
                    GroupTable =35
                    LayoutCachedLeft =11916
                    LayoutCachedTop =4056
                    LayoutCachedWidth =13512
                    LayoutCachedHeight =4308
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14736
                    Top =312
                    Width =696
                    Height =228
                    FontSize =8
                    TabIndex =56
                    ForeColor =5026082
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupId"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14736
                    LayoutCachedTop =312
                    LayoutCachedWidth =15432
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    GroupTable =64
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =14040
                            Top =312
                            Width =690
                            Height =228
                            FontSize =8
                            FontWeight =700
                            ForeColor =5026082
                            Name ="Label1201"
                            Caption ="CGroup:"
                            GroupTable =64
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =14040
                            LayoutCachedTop =312
                            LayoutCachedWidth =14730
                            LayoutCachedHeight =540
                            LayoutGroup =12
                            GroupTable =64
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
' See "Schedule_Form.cls"
