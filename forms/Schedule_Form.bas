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
                    Left =90
                    Width =2430
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
                    LayoutCachedLeft =90
                    LayoutCachedWidth =2520
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
                    Left =5625
                    Top =450
                    Width =1995
                    Height =255
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

                    LayoutCachedLeft =5625
                    LayoutCachedTop =450
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =705
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
                            Left =4860
                            Top =450
                            Width =705
                            Height =255
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label168"
                            Caption ="FSL Type:"
                            GroupTable =2
                            LayoutCachedLeft =4860
                            LayoutCachedTop =450
                            LayoutCachedWidth =5565
                            LayoutCachedHeight =705
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
                    Left =2760
                    Top =1785
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =16
                    Name ="SchOut"
                    ControlSource ="SchOut"
                    Format ="Short Date"
                    StatusBarText ="date schedule sent out"
                    GroupTable =19

                    LayoutCachedLeft =2760
                    LayoutCachedTop =1785
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =2055
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
                            Left =1875
                            Top =1785
                            Width =825
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label5"
                            Caption ="Sch Out"
                            GroupTable =19
                            LayoutCachedLeft =1875
                            LayoutCachedTop =1785
                            LayoutCachedWidth =2700
                            LayoutCachedHeight =2055
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
                    Left =5205
                    Top =1785
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =17
                    Name ="SchIn"
                    ControlSource ="SchIn"
                    Format ="Short Date"
                    StatusBarText ="date schedule in"
                    GroupTable =19

                    LayoutCachedLeft =5205
                    LayoutCachedTop =1785
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =2055
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
                            Left =3885
                            Top =1785
                            Width =1260
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label6"
                            Caption ="Sch IN"
                            GroupTable =19
                            LayoutCachedLeft =3885
                            LayoutCachedTop =1785
                            LayoutCachedWidth =5145
                            LayoutCachedHeight =2055
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
                    Left =8115
                    Top =1485
                    Width =255
                    TabIndex =19
                    Name ="SchTitleAssignee"
                    ControlSource ="SchTitleAssignee"
                    GroupTable =23

                    LayoutCachedLeft =8115
                    LayoutCachedTop =1485
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =1725
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
                            Left =6690
                            Top =1485
                            Width =1365
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label11"
                            Caption ="Assignee:"
                            GroupTable =23
                            LayoutCachedLeft =6690
                            LayoutCachedTop =1485
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =1725
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
                    Left =8115
                    Top =1785
                    Width =255
                    TabIndex =21
                    Name ="SchTitltoFA"
                    ControlSource ="SchTitltoFA"
                    GroupTable =23

                    LayoutCachedLeft =8115
                    LayoutCachedTop =1785
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =2025
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
                            Left =6690
                            Top =1785
                            Width =1365
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label14"
                            Caption ="Fleet Advantage:"
                            GroupTable =23
                            LayoutCachedLeft =6690
                            LayoutCachedTop =1785
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =2025
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
                    Left =8115
                    Top =2085
                    Width =255
                    TabIndex =22
                    Name ="SchTitletoClient"
                    ControlSource ="SchTitletoClient"
                    GroupTable =23

                    LayoutCachedLeft =8115
                    LayoutCachedTop =2085
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =2325
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
                            Left =6690
                            Top =2085
                            Width =1365
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Client:"
                            GroupTable =23
                            LayoutCachedLeft =6690
                            LayoutCachedTop =2085
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =2325
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
                    Left =8115
                    Top =2685
                    Width =255
                    TabIndex =25
                    Name ="LienHolder"
                    ControlSource ="LienHolder"
                    GroupTable =23

                    LayoutCachedLeft =8115
                    LayoutCachedTop =2685
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =2925
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
                            Left =6690
                            Top =2685
                            Width =1365
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label16"
                            Caption ="Lien Holder"
                            GroupTable =23
                            LayoutCachedLeft =6690
                            LayoutCachedTop =2685
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =2925
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
                    Left =9495
                    Top =2385
                    Width =4245
                    FontSize =9
                    TabIndex =24
                    Name ="LienHolderNm"
                    ControlSource ="LienHolderNm"
                    GroupTable =23

                    LayoutCachedLeft =9495
                    LayoutCachedTop =2385
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =2625
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
                            Left =8430
                            Top =2385
                            Width =1005
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Name:"
                            GroupTable =23
                            LayoutCachedLeft =8430
                            LayoutCachedTop =2385
                            LayoutCachedWidth =9435
                            LayoutCachedHeight =2625
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
                    Left =4050
                    Top =4095
                    Width =1170
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

                    LayoutCachedLeft =4050
                    LayoutCachedTop =4095
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =4395
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
                            Left =2340
                            Top =4095
                            Width =1650
                            Height =300
                            FontSize =9
                            FontWeight =700
                            Name ="Label19"
                            Caption ="Liability Amt:"
                            GroupTable =31
                            LayoutCachedLeft =2340
                            LayoutCachedTop =4095
                            LayoutCachedWidth =3990
                            LayoutCachedHeight =4395
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
                    Left =2340
                    Top =3405
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =33
                    Name ="SCHINREC"
                    ControlSource ="SCHINREC"
                    Format ="mm/dd/yy"
                    GroupTable =31

                    LayoutCachedLeft =2340
                    LayoutCachedTop =3405
                    LayoutCachedWidth =3270
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
                            Left =1470
                            Top =3405
                            Width =810
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Received:"
                            GroupTable =31
                            LayoutCachedLeft =1470
                            LayoutCachedTop =3405
                            LayoutCachedWidth =2280
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
                    Left =4050
                    Top =3405
                    Width =1170
                    Height =255
                    FontSize =9
                    TabIndex =34
                    Name ="SCHINSEXP"
                    ControlSource ="SCHINSEXP"
                    Format ="mm/dd/yy"
                    GroupTable =31

                    LayoutCachedLeft =4050
                    LayoutCachedTop =3405
                    LayoutCachedWidth =5220
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
                            Left =3330
                            Top =3405
                            Width =660
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =1643706
                            Name ="Label21"
                            Caption ="Expires:"
                            GroupTable =31
                            LayoutCachedLeft =3330
                            LayoutCachedTop =3405
                            LayoutCachedWidth =3990
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
                    Left =9495
                    Top =1485
                    Width =4245
                    Height =840
                    FontSize =9
                    TabIndex =20
                    Name ="SchTitleHolder"
                    ControlSource ="SchTitleHolder"
                    GroupTable =23

                    LayoutCachedLeft =9495
                    LayoutCachedTop =1485
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =2325
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
                            Left =8430
                            Top =1485
                            Width =1005
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label34"
                            Caption ="Title Holder:"
                            GroupTable =23
                            LayoutCachedLeft =8430
                            LayoutCachedTop =1485
                            LayoutCachedWidth =9435
                            LayoutCachedHeight =1725
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
                    Left =1215
                    Top =90
                    Width =3578
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

                    LayoutCachedLeft =1215
                    LayoutCachedTop =90
                    LayoutCachedWidth =4793
                    LayoutCachedHeight =390
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
                            Left =255
                            Top =90
                            Width =900
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label45"
                            Caption ="SchNo:"
                            GroupTable =2
                            LayoutCachedLeft =255
                            LayoutCachedTop =90
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =390
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
                    Left =8610
                    Top =90
                    Width =3990
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SchShortDesc"
                    ControlSource ="SchShortDesc"
                    GroupTable =2

                    LayoutCachedLeft =8610
                    LayoutCachedTop =90
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =390
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
                            Left =7680
                            Top =90
                            Width =870
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label48"
                            Caption ="Short Desc:"
                            GroupTable =2
                            LayoutCachedLeft =7680
                            LayoutCachedTop =90
                            LayoutCachedWidth =8550
                            LayoutCachedHeight =390
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
                    Left =14715
                    Top =105
                    Width =690
                    Height =225
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

                    LayoutCachedLeft =14715
                    LayoutCachedTop =105
                    LayoutCachedWidth =15405
                    LayoutCachedHeight =330
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =14025
                            Top =105
                            Width =690
                            Height =225
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
                            LayoutCachedLeft =14025
                            LayoutCachedTop =105
                            LayoutCachedWidth =14715
                            LayoutCachedHeight =330
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
                    Left =8190
                    Top =4035
                    Width =1605
                    Height =255
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

                    LayoutCachedLeft =8190
                    LayoutCachedTop =4035
                    LayoutCachedWidth =9795
                    LayoutCachedHeight =4290
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
                            Left =6675
                            Top =4035
                            Width =1455
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="FSLSCH_Label"
                            Caption ="FSL Sch:"
                            GroupTable =35
                            LayoutCachedLeft =6675
                            LayoutCachedTop =4035
                            LayoutCachedWidth =8130
                            LayoutCachedHeight =4290
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
                    Left =8190
                    Top =3450
                    Width =5265
                    Height =255
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

                    LayoutCachedLeft =8190
                    LayoutCachedTop =3450
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =3705
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
                            Left =6675
                            Top =3450
                            Width =1455
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label63"
                            Caption ="Billing Address:"
                            GroupTable =35
                            LayoutCachedLeft =6675
                            LayoutCachedTop =3450
                            LayoutCachedWidth =8130
                            LayoutCachedHeight =3705
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
                    Left =9015
                    Top =5490
                    Width =705
                    Height =285
                    TabIndex =44
                    Name ="BLCDACDT"
                    ControlSource ="BLCDACDT"
                    StatusBarText ="term starts with Accepted dt not BLCD"
                    GroupTable =35

                    LayoutCachedLeft =9015
                    LayoutCachedTop =5490
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =5775
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
                            Left =6675
                            Top =5490
                            Width =2287
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label80"
                            Caption ="Term Starts Acceptance date."
                            GroupTable =35
                            LayoutCachedLeft =6675
                            LayoutCachedTop =5490
                            LayoutCachedWidth =8962
                            LayoutCachedHeight =5775
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
                    Left =1485
                    Top =1485
                    Width =330
                    TabIndex =14
                    Name ="IntLease"
                    ControlSource ="IntLease"
                    GroupTable =19

                    LayoutCachedLeft =1485
                    LayoutCachedTop =1485
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1725
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
                            Left =405
                            Top =1485
                            Width =1020
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label87"
                            Caption ="Interim Lease"
                            GroupTable =19
                            LayoutCachedLeft =405
                            LayoutCachedTop =1485
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =1725
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
                    Left =11175
                    Top =5490
                    Width =2280
                    Height =285
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

                    LayoutCachedLeft =11175
                    LayoutCachedTop =5490
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =5775
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
                            Left =9855
                            Top =5490
                            Width =1260
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label57"
                            Caption ="Bank Client No:"
                            GroupTable =35
                            LayoutCachedLeft =9855
                            LayoutCachedTop =5490
                            LayoutCachedWidth =11115
                            LayoutCachedHeight =5775
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
                    Left =6675
                    Top =4350
                    Width =6780
                    Height =330
                    FontWeight =700
                    Name ="Label88"
                    Caption ="                            Customer Required Invoice \"Fields\""
                    GroupTable =35
                    LayoutCachedLeft =6675
                    LayoutCachedTop =4350
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =4680
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
                    Left =9855
                    Top =4740
                    Width =1260
                    Height =315
                    FontSize =9
                    TabIndex =40
                    Name ="InvFld1desc"
                    ControlSource ="InvFld1desc"
                    StatusBarText ="Lessee descriptor of number to appear on invoice"
                    GroupTable =35

                    LayoutCachedLeft =9855
                    LayoutCachedTop =4740
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =5055
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
                    Left =11865
                    Top =4740
                    Width =1590
                    Height =315
                    FontSize =9
                    TabIndex =41
                    Name ="InvFld1"
                    ControlSource ="InvFld1"
                    StatusBarText ="number itself"
                    GroupTable =35

                    LayoutCachedLeft =11865
                    LayoutCachedTop =4740
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =5055
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
                            Left =11175
                            Top =4740
                            Width =630
                            Height =315
                            FontSize =9
                            FontWeight =700
                            Name ="Label94"
                            Caption ="Value:"
                            GroupTable =35
                            LayoutCachedLeft =11175
                            LayoutCachedTop =4740
                            LayoutCachedWidth =11805
                            LayoutCachedHeight =5055
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
                    Left =9855
                    Top =5115
                    Width =1260
                    Height =315
                    FontSize =9
                    TabIndex =42
                    Name ="InvFld2desc"
                    ControlSource ="InvFld2desc"
                    StatusBarText ="Lessee descriptor of number to appear on invoice"
                    GroupTable =35

                    LayoutCachedLeft =9855
                    LayoutCachedTop =5115
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =5430
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
                    Left =11865
                    Top =5115
                    Width =1590
                    Height =315
                    FontSize =9
                    TabIndex =43
                    Name ="Text97"
                    ControlSource ="InvFld2"
                    StatusBarText ="number itself"
                    GroupTable =35

                    LayoutCachedLeft =11865
                    LayoutCachedTop =5115
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =5430
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
                            Left =11175
                            Top =5115
                            Width =630
                            Height =315
                            FontSize =9
                            FontWeight =700
                            Name ="Label98"
                            Caption ="Value:"
                            GroupTable =35
                            LayoutCachedLeft =11175
                            LayoutCachedTop =5115
                            LayoutCachedWidth =11805
                            LayoutCachedHeight =5430
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
                    Left =8115
                    Top =2385
                    Width =255
                    TabIndex =23
                    Name ="SchTitleToTrust"
                    ControlSource ="SchTitleToTrust"
                    GroupTable =23

                    LayoutCachedLeft =8115
                    LayoutCachedTop =2385
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =2625
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
                            Left =6690
                            Top =2385
                            Width =1365
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label101"
                            Caption ="FA Title Trust:"
                            GroupTable =23
                            LayoutCachedLeft =6690
                            LayoutCachedTop =2385
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =2625
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
                    Left =4050
                    Top =3720
                    Width =1170
                    Height =315
                    FontSize =9
                    TabIndex =35
                    Name ="SchSLVRec"
                    ControlSource ="SchSLVRec"
                    Format ="mm/dd/yy"
                    StatusBarText ="SLV Tables received"
                    GroupTable =31

                    LayoutCachedLeft =4050
                    LayoutCachedTop =3720
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =4035
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
                            Left =1470
                            Top =3720
                            Width =2520
                            Height =315
                            FontSize =9
                            FontWeight =700
                            Name ="Label109"
                            Caption ="Stip Loss Table Received:"
                            GroupTable =31
                            LayoutCachedLeft =1470
                            LayoutCachedTop =3720
                            LayoutCachedWidth =3990
                            LayoutCachedHeight =4035
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
                    Left =5205
                    Top =1485
                    Width =1065
                    FontSize =9
                    TabIndex =15
                    Name ="SchedDt"
                    ControlSource ="SchedDt"
                    Format ="Short Date"
                    GroupTable =19

                    LayoutCachedLeft =5205
                    LayoutCachedTop =1485
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =1725
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
                            Left =3885
                            Top =1485
                            Width =1260
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label115"
                            Caption ="Sch Date"
                            GroupTable =19
                            LayoutCachedLeft =3885
                            LayoutCachedTop =1485
                            LayoutCachedWidth =5145
                            LayoutCachedHeight =1725
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
                                    Left =6270
                                    Top =7470
                                    Height =270
                                    FontSize =9
                                    TabIndex =1
                                    Name ="SchFEcoRec"
                                    ControlSource ="SchFEcoRec"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="date final economics received"
                                    GroupTable =55

                                    LayoutCachedLeft =6270
                                    LayoutCachedTop =7470
                                    LayoutCachedWidth =7710
                                    LayoutCachedHeight =7740
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
                                            Left =4125
                                            Top =7470
                                            Width =2085
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label145"
                                            Caption ="Synd Economics Received:"
                                            GroupTable =55
                                            LayoutCachedLeft =4125
                                            LayoutCachedTop =7470
                                            LayoutCachedWidth =6210
                                            LayoutCachedHeight =7740
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
                                    Left =6270
                                    Top =7800
                                    Height =270
                                    FontSize =9
                                    TabIndex =2
                                    Name ="SchFEcoA"
                                    ControlSource ="SchFEcoA"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="date final economics apporved"
                                    GroupTable =55

                                    LayoutCachedLeft =6270
                                    LayoutCachedTop =7800
                                    LayoutCachedWidth =7710
                                    LayoutCachedHeight =8070
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
                                            Left =4125
                                            Top =7800
                                            Width =2085
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label146"
                                            Caption ="Synd Economics Approved:"
                                            GroupTable =55
                                            LayoutCachedLeft =4125
                                            LayoutCachedTop =7800
                                            LayoutCachedWidth =6210
                                            LayoutCachedHeight =8070
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
                                    Left =6270
                                    Top =8130
                                    Height =270
                                    FontSize =9
                                    TabIndex =3
                                    Name ="Text151"
                                    ControlSource ="SchFPkgO"
                                    StatusBarText ="date Funding package sent to bank"
                                    GroupTable =55

                                    LayoutCachedLeft =6270
                                    LayoutCachedTop =8130
                                    LayoutCachedWidth =7710
                                    LayoutCachedHeight =8400
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
                                            Left =4125
                                            Top =8130
                                            Width =2085
                                            Height =270
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label152"
                                            Caption ="Perm Funding Pkg. Out:"
                                            GroupTable =55
                                            LayoutCachedLeft =4125
                                            LayoutCachedTop =8130
                                            LayoutCachedWidth =6210
                                            LayoutCachedHeight =8400
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
                                    Left =13200
                                    Top =7560
                                    Width =1590
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

                                    LayoutCachedLeft =13200
                                    LayoutCachedTop =7560
                                    LayoutCachedWidth =14790
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
                                            Left =8655
                                            Top =7560
                                            Width =4485
                                            Height =255
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label81"
                                            Caption ="Per Unit CurrRent:(Sliver Rent Calculation):"
                                            GroupTable =63
                                            LayoutCachedLeft =8655
                                            LayoutCachedTop =7560
                                            LayoutCachedWidth =13140
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
                                    Left =13200
                                    Top =7875
                                    Width =1590
                                    Height =255
                                    FontSize =9
                                    TabIndex =6
                                    Name ="SchCRExp"
                                    ControlSource ="SchCRExp"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="current rent expiry"
                                    GroupTable =63

                                    LayoutCachedLeft =13200
                                    LayoutCachedTop =7875
                                    LayoutCachedWidth =14790
                                    LayoutCachedHeight =8130
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
                                            Left =11190
                                            Top =7875
                                            Width =1950
                                            Height =255
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label82"
                                            Caption ="Old Lease Expiry:"
                                            GroupTable =63
                                            LayoutCachedLeft =11190
                                            LayoutCachedTop =7875
                                            LayoutCachedWidth =13140
                                            LayoutCachedHeight =8130
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
                                    Left =8655
                                    Top =7875
                                    Width =2460
                                    Height =255
                                    Name ="EmptyCell1196"
                                    GroupTable =63
                                    RightPadding =0
                                    LayoutCachedLeft =8655
                                    LayoutCachedTop =7875
                                    LayoutCachedWidth =11115
                                    LayoutCachedHeight =8130
                                    RowStart =1
                                    RowEnd =1
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11115
                                    Top =7875
                                    Width =0
                                    Height =255
                                    Name ="EmptyCell1197"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11115
                                    LayoutCachedTop =7875
                                    LayoutCachedWidth =11115
                                    LayoutCachedHeight =8130
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11115
                                    Top =7875
                                    Width =45
                                    Height =255
                                    Name ="EmptyCell1198"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11115
                                    LayoutCachedTop =7875
                                    LayoutCachedWidth =11160
                                    LayoutCachedHeight =8130
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =11
                                    GroupTable =63
                                End
                                Begin EmptyCell
                                    Left =11160
                                    Top =7875
                                    Width =0
                                    Height =255
                                    Name ="EmptyCell1199"
                                    GroupTable =63
                                    LeftPadding =0
                                    RightPadding =0
                                    LayoutCachedLeft =11160
                                    LayoutCachedTop =7875
                                    LayoutCachedWidth =11160
                                    LayoutCachedHeight =8130
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
                    Left =3210
                    Top =450
                    Width =570
                    Height =255
                    TabIndex =4
                    Name ="SchAH"
                    ControlSource ="SchAH"
                    StatusBarText ="Autohauler schedule"
                    GroupTable =2

                    LayoutCachedLeft =3210
                    LayoutCachedTop =450
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =705
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
                            Left =1890
                            Top =450
                            Width =1260
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label121"
                            Caption ="Autohaulers:"
                            GroupTable =2
                            LayoutCachedLeft =1890
                            LayoutCachedTop =450
                            LayoutCachedWidth =3150
                            LayoutCachedHeight =705
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
                    Left =13185
                    Top =3765
                    Width =270
                    Height =210
                    TabIndex =38
                    Name ="SchFABill"
                    ControlSource ="SchFABill"
                    StatusBarText ="FA Bills an assigned schedule on behalf of bank"
                    GroupTable =35

                    LayoutCachedLeft =13185
                    LayoutCachedTop =3765
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =3975
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
                            Left =11175
                            Top =3765
                            Width =1950
                            Height =210
                            FontSize =8
                            FontWeight =700
                            Name ="Label125"
                            Caption ="FA Bills after Assignment:"
                            GroupTable =35
                            LayoutCachedLeft =11175
                            LayoutCachedTop =3765
                            LayoutCachedWidth =13125
                            LayoutCachedHeight =3975
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
                    Left =1215
                    Top =450
                    Width =611
                    Height =255
                    TabIndex =3
                    Name ="SchExSyn"
                    ControlSource ="SchExSyn"
                    StatusBarText ="schedule is not to be syndicated added 5-17-13"
                    GroupTable =2

                    LayoutCachedLeft =1215
                    LayoutCachedTop =450
                    LayoutCachedWidth =1826
                    LayoutCachedHeight =705
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
                            Left =255
                            Top =450
                            Width =900
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label126"
                            Caption ="Do Not Syn: "
                            GroupTable =2
                            LayoutCachedLeft =255
                            LayoutCachedTop =450
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =705
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
                    Left =5625
                    Top =90
                    Width =1995
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

                    LayoutCachedLeft =5625
                    LayoutCachedTop =90
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =390
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
                            Left =4860
                            Top =90
                            Width =705
                            Height =300
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label110"
                            Caption ="FAID:"
                            GroupTable =2
                            LayoutCachedLeft =4860
                            LayoutCachedTop =90
                            LayoutCachedWidth =5565
                            LayoutCachedHeight =390
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
                    Left =13335
                    Top =105
                    Width =690
                    Height =225
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

                    LayoutCachedLeft =13335
                    LayoutCachedTop =105
                    LayoutCachedWidth =14025
                    LayoutCachedHeight =330
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12645
                            Top =105
                            Width =690
                            Height =225
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
                            LayoutCachedLeft =12645
                            LayoutCachedTop =105
                            LayoutCachedWidth =13335
                            LayoutCachedHeight =330
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
                    Left =16095
                    Top =105
                    Width =690
                    Height =225
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

                    LayoutCachedLeft =16095
                    LayoutCachedTop =105
                    LayoutCachedWidth =16785
                    LayoutCachedHeight =330
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GroupTable =13
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =15405
                            Top =105
                            Width =690
                            Height =225
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
                            LayoutCachedLeft =15405
                            LayoutCachedTop =105
                            LayoutCachedWidth =16095
                            LayoutCachedHeight =330
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
                    Left =4440
                    Top =450
                    Width =353
                    Height =255
                    TabIndex =5
                    BorderColor =8210719
                    Name ="SchFSL"
                    ControlSource ="SchFSL"
                    StatusBarText ="FSL Lease"
                    GroupTable =2

                    LayoutCachedLeft =4440
                    LayoutCachedTop =450
                    LayoutCachedWidth =4793
                    LayoutCachedHeight =705
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
                            Left =3840
                            Top =450
                            Width =540
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label46"
                            Caption ="FSL:"
                            GroupTable =2
                            LayoutCachedLeft =3840
                            LayoutCachedTop =450
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =705
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
                    Top =2700
                    Width =690
                    Name ="EmptyCell284"
                    GroupTable =16
                    LayoutCachedLeft =420
                    LayoutCachedTop =2700
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =2940
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =420
                    Top =2400
                    Width =690
                    Name ="EmptyCell287"
                    GroupTable =16
                    LayoutCachedLeft =420
                    LayoutCachedTop =2400
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =2640
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
                    Left =8430
                    Top =1785
                    Width =1005
                    Name ="EmptyCell392"
                    GroupTable =23
                    LayoutCachedLeft =8430
                    LayoutCachedTop =1785
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =2025
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =8430
                    Top =2085
                    Width =1005
                    Name ="EmptyCell393"
                    GroupTable =23
                    LayoutCachedLeft =8430
                    LayoutCachedTop =2085
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =2325
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =8430
                    Top =2685
                    Width =1005
                    Name ="EmptyCell401"
                    GroupTable =23
                    LayoutCachedLeft =8430
                    LayoutCachedTop =2685
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =2925
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =9495
                    Top =2685
                    Width =4245
                    Name ="EmptyCell402"
                    GroupTable =23
                    LayoutCachedLeft =9495
                    LayoutCachedTop =2685
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =2925
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
                    Left =1470
                    Top =4095
                    Width =810
                    Height =300
                    Name ="EmptyCell437"
                    GroupTable =31
                    LayoutCachedLeft =1470
                    LayoutCachedTop =4095
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =4395
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
                    Left =6675
                    Top =3765
                    Width =4440
                    Height =210
                    Name ="EmptyCell561"
                    GroupTable =35
                    LayoutCachedLeft =6675
                    LayoutCachedTop =3765
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =3975
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =4
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =6675
                    Top =4740
                    Width =3045
                    Height =315
                    FontSize =9
                    FontWeight =700
                    Name ="Label89"
                    Caption ="Description/Label 1:"
                    GroupTable =35
                    LayoutCachedLeft =6675
                    LayoutCachedTop =4740
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =5055
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
                    Left =6675
                    Top =5115
                    Width =3045
                    Height =315
                    FontSize =9
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Description/Label 2:"
                    GroupTable =35
                    LayoutCachedLeft =6675
                    LayoutCachedTop =5115
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =5430
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
                    Left =9780
                    Top =4740
                    Width =15
                    Height =315
                    Name ="EmptyCell604"
                    GroupTable =35
                    LayoutCachedLeft =9780
                    LayoutCachedTop =4740
                    LayoutCachedWidth =9795
                    LayoutCachedHeight =5055
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin EmptyCell
                    Left =9780
                    Top =5115
                    Width =15
                    Height =315
                    Name ="EmptyCell605"
                    GroupTable =35
                    LayoutCachedLeft =9780
                    LayoutCachedTop =5115
                    LayoutCachedWidth =9795
                    LayoutCachedHeight =5430
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
                    Left =9780
                    Top =5490
                    Width =15
                    Height =285
                    Name ="EmptyCell675"
                    GroupTable =35
                    LayoutCachedLeft =9780
                    LayoutCachedTop =5490
                    LayoutCachedWidth =9795
                    LayoutCachedHeight =5775
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
                    Left =16095
                    Top =330
                    Width =690
                    Height =225
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

                    LayoutCachedLeft =16095
                    LayoutCachedTop =330
                    LayoutCachedWidth =16785
                    LayoutCachedHeight =555
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =8
                    GroupTable =57
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =15405
                            Top =330
                            Width =690
                            Height =225
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
                            LayoutCachedLeft =15405
                            LayoutCachedTop =330
                            LayoutCachedWidth =16095
                            LayoutCachedHeight =555
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
                    Left =2760
                    Top =1170
                    Width =1065
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="ApprovalRequest"
                    ControlSource ="ApprovalRequest"
                    Format ="Short Date"
                    StatusBarText ="Internal Approval Request Date"
                    GroupTable =19

                    LayoutCachedLeft =2760
                    LayoutCachedTop =1170
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =1425
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
                            Left =405
                            Top =1170
                            Width =2295
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label855"
                            Caption ="Schedule Draft Out "
                            GroupTable =19
                            LayoutCachedLeft =405
                            LayoutCachedTop =1170
                            LayoutCachedWidth =2700
                            LayoutCachedHeight =1425
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
                    Left =5205
                    Top =1170
                    Width =1065
                    Height =255
                    FontSize =9
                    TabIndex =13
                    Name ="ApprovalReply"
                    ControlSource ="ApprovalReply"
                    Format ="Short Date"
                    StatusBarText ="Internal Approval REPLY Date"
                    GroupTable =19

                    LayoutCachedLeft =5205
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =1425
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
                            Left =3885
                            Top =1170
                            Width =1260
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label857"
                            Caption ="Draft Approved "
                            GroupTable =19
                            LayoutCachedLeft =3885
                            LayoutCachedTop =1170
                            LayoutCachedWidth =5145
                            LayoutCachedHeight =1425
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
                    Left =1170
                    Top =2100
                    Width =5085
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

                    LayoutCachedLeft =1170
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6255
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
                    Left =1170
                    Top =2400
                    Width =5085
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

                    LayoutCachedLeft =1170
                    LayoutCachedTop =2400
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =2640
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
                    Left =1170
                    Top =2700
                    Width =5085
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

                    LayoutCachedLeft =1170
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =2940
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
                    Left =7680
                    Top =450
                    Width =870
                    Height =255
                    Name ="EmptyCell1088"
                    GroupTable =2
                    LayoutCachedLeft =7680
                    LayoutCachedTop =450
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8610
                    Top =450
                    Width =3990
                    Height =255
                    Name ="EmptyCell1089"
                    GroupTable =2
                    LayoutCachedLeft =8610
                    LayoutCachedTop =450
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6315
                    Top =2400
                    Width =90
                    Name ="EmptyCell1098"
                    GroupTable =16
                    LayoutCachedLeft =6315
                    LayoutCachedTop =2400
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =2640
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =6315
                    Top =2700
                    Width =90
                    Name ="EmptyCell1099"
                    GroupTable =16
                    LayoutCachedLeft =6315
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =2940
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =9
                    GroupTable =16
                End
                Begin EmptyCell
                    Left =6315
                    Top =2100
                    Width =90
                    Name ="EmptyCell1100"
                    GroupTable =16
                    LayoutCachedLeft =6315
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6405
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
                    Left =1485
                    Top =1785
                    Width =330
                    Height =270
                    Name ="EmptyCell1182"
                    GroupTable =19
                    LayoutCachedLeft =1485
                    LayoutCachedTop =1785
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =2055
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =405
                    Top =1785
                    Width =1020
                    Height =270
                    Name ="EmptyCell1183"
                    GroupTable =19
                    LayoutCachedLeft =405
                    LayoutCachedTop =1785
                    LayoutCachedWidth =1425
                    LayoutCachedHeight =2055
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =2760
                    Top =1485
                    Width =1065
                    Name ="EmptyCell1190"
                    GroupTable =19
                    LayoutCachedLeft =2760
                    LayoutCachedTop =1485
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =1725
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =1875
                    Top =1485
                    Width =825
                    Name ="EmptyCell1191"
                    GroupTable =19
                    LayoutCachedLeft =1875
                    LayoutCachedTop =1485
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1725
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =19
                End
                Begin EmptyCell
                    Left =9855
                    Top =4035
                    Width =1950
                    Height =255
                    Name ="EmptyCell1194"
                    GroupTable =35
                    LayoutCachedLeft =9855
                    LayoutCachedTop =4035
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =4290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =6
                    GroupTable =35
                End
                Begin EmptyCell
                    Left =11865
                    Top =4035
                    Width =1590
                    Height =255
                    Name ="EmptyCell1195"
                    GroupTable =35
                    LayoutCachedLeft =11865
                    LayoutCachedTop =4035
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =4290
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
                    Left =14730
                    Top =315
                    Width =690
                    Height =225
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

                    LayoutCachedLeft =14730
                    LayoutCachedTop =315
                    LayoutCachedWidth =15420
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =12
                    GroupTable =64
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =14040
                            Top =315
                            Width =690
                            Height =225
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
                            LayoutCachedTop =315
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
