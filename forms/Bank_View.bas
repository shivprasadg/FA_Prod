Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =23640
    DatasheetFontHeight =10
    ItemSuffix =612
    Left =5745
    Top =4320
    Right =29640
    Bottom =17700
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Banks].[BankId], [Banks].[BankShortName]"
    RecSrcDt = Begin
        0xf42f3532335de540
    End
    RecordSource ="Banks"
    Caption ="Bank_View"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
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
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =615
            Name ="FormHeader"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Top =15
                    Width =10920
                    Height =480
                    ColumnOrder =0
                    FontSize =18
                    Name ="ClientName"
                    ControlSource ="BankName"
                    FontName ="Segoe UI"

                    LayoutCachedTop =15
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =495
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =14205
                    Top =120
                    Width =1410
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Command382"
                    Caption ="Add Bank"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command382\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Action "
                                "Name=\"GoToRecord\"><Argument Nam"
                        End
                        Begin
                            Comment ="_AXL:e=\"Record\">New</Argument></Action><ConditionalBlock><If><Condition>[Macro"
                                "Error]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argument Nam"
                                "e=\"Message\">=[MacroError].[Description]</Argument></Action></Statements></If><"
                                "/ConditionalBlock><"
                        End
                        Begin
                            Comment ="_AXL:/Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14205
                    LayoutCachedTop =120
                    LayoutCachedWidth =15615
                    LayoutCachedHeight =408
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =15660
                    Top =120
                    Width =1560
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="AllBankList"
                    Caption ="Bank List Report"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15660
                    LayoutCachedTop =120
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =408
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =238
                    Left =17280
                    Top =120
                    Width =1290
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="AllBankExcel"
                    Caption ="Bank List Excel"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17280
                    LayoutCachedTop =120
                    LayoutCachedWidth =18570
                    LayoutCachedHeight =408
                    ForeThemeColorIndex =6
                    ForeShade =75.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =12255
                    Top =180
                    Width =1770
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =4
                    Name ="BankGroupID"
                    ControlSource ="BankGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientGroups.ClientGroupId, ClientGroups.ClientGroupName\015\012FROM Clie"
                        "ntGroups\015\012GROUP BY ClientGroups.ClientGroupId, ClientGroups.ClientGroupNam"
                        "e\015\012ORDER BY ClientGroups.ClientGroupName;\015\012"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    FontName ="Calibri"

                    LayoutCachedLeft =12255
                    LayoutCachedTop =180
                    LayoutCachedWidth =14025
                    LayoutCachedHeight =435
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10980
                            Top =180
                            Width =1215
                            Height =255
                            Name ="Label331"
                            Caption ="BankGroup:"
                            FontName ="Calibri"
                            LayoutCachedLeft =10980
                            LayoutCachedTop =180
                            LayoutCachedWidth =12195
                            LayoutCachedHeight =435
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =19875
                    Top =165
                    Width =750
                    Height =228
                    TabIndex =5
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =19875
                    LayoutCachedTop =165
                    LayoutCachedWidth =20625
                    LayoutCachedHeight =393
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
                    OverlapFlags =215
                    Left =19320
                    Top =165
                    Width =555
                    Height =228
                    TabIndex =6
                    Name ="GoHome"
                    Caption ="Home"
                    FontName ="Calibri"
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

                    LayoutCachedLeft =19320
                    LayoutCachedTop =165
                    LayoutCachedWidth =19875
                    LayoutCachedHeight =393
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =12780
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =2325
                    Top =2040
                    Width =21315
                    Height =10695
                    FontSize =10
                    TabIndex =23
                    Name ="TabCtl155"
                    FontName ="Calibri"
                    VerticalAnchor =2

                    LayoutCachedLeft =2325
                    LayoutCachedTop =2040
                    LayoutCachedWidth =23640
                    LayoutCachedHeight =12735
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =2460
                            Top =2475
                            Width =21045
                            Height =10125
                            BorderColor =0
                            Name ="Contacts"
                            Caption ="Contacts"
                            GridlineColor =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12600
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =4395
                                    Top =3555
                                    Width =1350
                                    Height =345
                                    FontWeight =700
                                    ForeColor =-2147483615
                                    Name ="AddContactBtnBnk"
                                    Caption ="Add a Contact"
                                    OnClickEmMacro = Begin
                                        Version =196611
                                        ColumnsShown =12
                                        Begin
                                            Action ="OpenForm"
                                            Argument ="bankContact_view"
                                            Argument ="0"
                                            Argument =""
                                            Argument =""
                                            Argument ="0"
                                            Argument ="0"
                                        End
                                        Begin
                                            Action ="SetValue"
                                            Argument ="[Forms]![bankcontact_view]![clientref]"
                                            Argument ="[Forms]![bank_view]![clientid]"
                                        End
                                        Begin
                                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                                "nterfaceMacro For=\"AddContactBtnBnk\" xmlns=\"http://schemas.microsoft.com/offi"
                                                "ce/accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><A"
                                                "rgument Name=\"FormName\">bankCon"
                                        End
                                        Begin
                                            Comment ="_AXL:tact_view</Argument><Argument Name=\"DataMode\">Add</Argument></Action><Act"
                                                "ion Name=\"SetValue\"><Argument Name=\"Item\">[Forms]![bankcontact_view]![client"
                                                "ref]</Argument><Argument Name=\"Expression\">[Forms]![bank_view]![clientid]</Arg"
                                                "ument></Action></Stat"
                                        End
                                        Begin
                                            Comment ="_AXL:ements></UserInterfaceMacro>"
                                        End
                                    End

                                    LayoutCachedLeft =4395
                                    LayoutCachedTop =3555
                                    LayoutCachedWidth =5745
                                    LayoutCachedHeight =3900
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =6015
                                    Top =3555
                                    Height =345
                                    FontWeight =700
                                    TabIndex =1
                                    ForeColor =-2147483615
                                    Name ="EditContactsBtnBnk"
                                    Caption ="Edit Contacts"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =6015
                                    LayoutCachedTop =3555
                                    LayoutCachedWidth =7455
                                    LayoutCachedHeight =3900
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =7830
                                    Top =3555
                                    Height =345
                                    FontWeight =700
                                    TabIndex =2
                                    ForeColor =-2147483615
                                    Name ="reqqBnk"
                                    Caption ="Refresh"
                                    OnClickEmMacro = Begin
                                        Version =196611
                                        ColumnsShown =8
                                        Begin
                                            Action ="Requery"
                                            Argument ="Contacts_subBnk"
                                        End
                                        Begin
                                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                                "nterfaceMacro For=\"reqqBnk\" xmlns=\"http://schemas.microsoft.com/office/access"
                                                "services/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Na"
                                                "me=\"ControlName\">Contacts_subBn"
                                        End
                                        Begin
                                            Comment ="_AXL:k</Argument></Action></Statements></UserInterfaceMacro>"
                                        End
                                    End

                                    LayoutCachedLeft =7830
                                    LayoutCachedTop =3555
                                    LayoutCachedWidth =9270
                                    LayoutCachedHeight =3900
                                    Overlaps =1
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =2520
                                    Top =2595
                                    Width =20820
                                    Height =9165
                                    TabIndex =3
                                    Name ="Contacts_subBnk"
                                    SourceObject ="Form.BankContact_sub"
                                    LinkChildFields ="ClientId"
                                    LinkMasterFields ="BankID"

                                    LayoutCachedLeft =2520
                                    LayoutCachedTop =2595
                                    LayoutCachedWidth =23340
                                    LayoutCachedHeight =11760
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =2460
                            Top =2475
                            Width =21045
                            Height =10125
                            BorderColor =0
                            Name ="Assigned Schedules"
                            EventProcPrefix ="Assigned_Schedules"
                            GridlineColor =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12600
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =0
                                    Left =2685
                                    Top =2895
                                    Width =20400
                                    Height =9645
                                    Name ="BankSchSub"
                                    SourceObject ="Form.BankSchSub"
                                    LinkChildFields ="Assignee_Bank"
                                    LinkMasterFields ="BankShortName"

                                    LayoutCachedLeft =2685
                                    LayoutCachedTop =2895
                                    LayoutCachedWidth =23085
                                    LayoutCachedHeight =12540
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =2685
                                    Top =2475
                                    Width =3840
                                    Height =345
                                    FontSize =14
                                    FontWeight =700
                                    Name ="Label582"
                                    Caption ="Assigned Schedules"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =2685
                                    LayoutCachedTop =2475
                                    LayoutCachedWidth =6525
                                    LayoutCachedHeight =2820
                                    ForeThemeColorIndex =2
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =2459
                            Top =2475
                            Width =21046
                            Height =10130
                            BorderColor =0
                            Name ="Bank Rate && Wire Info | Title Request"
                            EventProcPrefix ="Bank_Rate____Wire_Info___Title_Request"
                            GridlineColor =0
                            LayoutCachedLeft =2459
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12605
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2460
                                    Top =2640
                                    Width =9120
                                    Height =6180
                                    Name ="BankRates_sub"
                                    SourceObject ="Form.BankRates_sub"
                                    LinkChildFields ="BankID"
                                    LinkMasterFields ="BankID"

                                    LayoutCachedLeft =2460
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =11580
                                    LayoutCachedHeight =8820
                                End
                                Begin ListBox
                                    OverlapFlags =247
                                    MultiSelect =2
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =12135
                                    Top =4065
                                    Width =6945
                                    Height =8415
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =1
                                    BorderColor =14211288
                                    Name ="Units_List"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT UnitTitleHistoryLastEntry.AssetID, UnitTitleHistoryLastEntry.Vin, UnitTit"
                                        "leHistoryLastEntry.UnitNum, UnitTitleHistoryLastEntry.uStatus, UnitTitleHistoryL"
                                        "astEntry.Client, UnitTitleHistoryLastEntry.Schedule, UnitTitleHistoryLastEntry.G"
                                        "roupID FROM UnitTitleHistoryLastEntry WHERE (((UnitTitleHistoryLastEntry.EntryRe"
                                        "asonId)=4) AND ((UnitTitleHistoryLastEntry.BankID)=[Forms]![Bank_View]![BankID])"
                                        ") OR (((UnitTitleHistoryLastEntry.EntryReasonId)=25) AND ((UnitTitleHistoryLastE"
                                        "ntry.BankID)=[Forms]![Bank_View]![BankID])) OR (((UnitTitleHistoryLastEntry.Entr"
                                        "yReasonId)=1) AND ((UnitTitleHistoryLastEntry.BankID)=[Forms]![Bank_View]![BankI"
                                        "D])) ORDER BY UnitTitleHistoryLastEntry.AssetID;"
                                    ColumnWidths ="720;1872;720;288;1536;1872"
                                    OnClick ="[Event Procedure]"
                                    VerticalAnchor =2
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =12135
                                    LayoutCachedTop =4065
                                    LayoutCachedWidth =19080
                                    LayoutCachedHeight =12480
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =2
                                    ForeShade =50.0
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =12120
                                    Top =3570
                                    Width =720
                                    FontSize =9
                                    TabIndex =2
                                    BackColor =13434879
                                    BorderColor =12632256
                                    Name ="SearchAssetID"
                                    FontName ="Calibri"
                                    OnLostFocus ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"

                                    LayoutCachedLeft =12120
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =3810
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =13350
                                    Top =3240
                                    Width =2595
                                    Height =285
                                    FontSize =11
                                    ForeColor =-2147483615
                                    Name ="GroupInstructions"
                                    Caption ="▲ Select Sch | Group ▲"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =13350
                                    LayoutCachedTop =3240
                                    LayoutCachedWidth =15945
                                    LayoutCachedHeight =3525
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    Left =14700
                                    Top =3255
                                    Width =2520
                                    Height =255
                                    FontSize =10
                                    TabIndex =3
                                    Name ="GetSelectedUnits"
                                    Caption ="▼Selected (1) Unit ►"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =14700
                                    LayoutCachedTop =3255
                                    LayoutCachedWidth =17220
                                    LayoutCachedHeight =3510
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =8210719
                                    BorderThemeColorIndex =1
                                    HoverColor =10319446
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =30
                                    QuickStyleMask =-117
                                    Overlaps =1
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =12840
                                    Top =3570
                                    Width =1872
                                    FontSize =9
                                    TabIndex =4
                                    BackColor =13434879
                                    BorderColor =12632256
                                    Name ="SearchVIN"
                                    FontName ="Calibri"
                                    OnLostFocus ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"

                                    LayoutCachedLeft =12840
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =14712
                                    LayoutCachedHeight =3810
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =13200
                                    Top =3255
                                    Width =1470
                                    Height =255
                                    FontSize =10
                                    TabIndex =5
                                    Name ="GetAll"
                                    Caption ="▼All 323"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =13200
                                    LayoutCachedTop =3255
                                    LayoutCachedWidth =14670
                                    LayoutCachedHeight =3510
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    BackColor =4754549
                                    BorderThemeColorIndex =1
                                    HoverColor =5026082
                                    PressedThemeColorIndex =6
                                    PressedShade =75.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =32
                                    QuickStyleMask =-117
                                    Overlaps =1
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =14712
                                    Top =3570
                                    Width =720
                                    FontSize =9
                                    TabIndex =6
                                    BackColor =13434879
                                    BorderColor =12632256
                                    Name ="SearchUnitNum"
                                    FontName ="Calibri"
                                    OnLostFocus ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"

                                    LayoutCachedLeft =14712
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =15432
                                    LayoutCachedHeight =3810
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =15717
                                    Top =3570
                                    Width =1537
                                    FontSize =9
                                    TabIndex =7
                                    BackColor =13434879
                                    BorderColor =12632256
                                    Name ="SearchClient"
                                    FontName ="Calibri"
                                    OnLostFocus ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"

                                    LayoutCachedLeft =15717
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =17254
                                    LayoutCachedHeight =3810
                                End
                                Begin ToggleButton
                                    OverlapFlags =223
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =12840
                                    Top =3810
                                    Width =1872
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =8
                                    ForeColor =8210719
                                    Name ="ToggleSortVIN"
                                    DefaultValue ="0"
                                    Caption ="↓↑ VIN"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =12840
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =14712
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin ToggleButton
                                    OverlapFlags =215
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =12120
                                    Top =3810
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =9
                                    ForeColor =8210719
                                    Name ="ToggleSortAssetID"
                                    DefaultValue ="0"
                                    Caption ="AssetID"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =12120
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin ToggleButton
                                    OverlapFlags =223
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =14712
                                    Top =3810
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =10
                                    ForeColor =8210719
                                    Name ="ToggleSortUnitNum"
                                    DefaultValue ="0"
                                    Caption ="Unit #"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =14712
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =15432
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin ToggleButton
                                    OverlapFlags =223
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =15717
                                    Top =3810
                                    Width =1537
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =11
                                    ForeColor =8210719
                                    Name ="ToggleSortClient"
                                    DefaultValue ="0"
                                    Caption ="Client"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =15717
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =17254
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin ToggleButton
                                    OverlapFlags =215
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =12135
                                    Top =3255
                                    Width =1005
                                    Height =255
                                    FontWeight =400
                                    TabIndex =12
                                    ForeColor =8355711
                                    Name ="PickModeToggle"
                                    DefaultValue ="=1"
                                    Caption ="▼Pick On"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =0

                                    LayoutCachedLeft =12135
                                    LayoutCachedTop =3255
                                    LayoutCachedWidth =13140
                                    LayoutCachedHeight =3510
                                    ForeTint =50.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15060409
                                    BackTint =40.0
                                    OldBorderStyle =1
                                    BorderThemeColorIndex =1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverThemeColorIndex =8
                                    HoverTint =100.0
                                    HoverShade =75.0
                                    PressedColor =0
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeThemeColorIndex =1
                                    HoverForeTint =100.0
                                    QuickStyle =30
                                    QuickStyleMask =-629
                                    Overlaps =1
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =22405
                                    Top =3570
                                    Width =839
                                    Height =299
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =13
                                    BorderColor =1643706
                                    Name ="UnitSelected"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =22405
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =23244
                                    LayoutCachedHeight =3869
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OldBorderStyle =1
                                            OverlapFlags =223
                                            TextAlign =2
                                            Left =22380
                                            Top =3360
                                            Width =900
                                            Height =540
                                            BackColor =1643706
                                            BorderColor =1643706
                                            ForeColor =16777215
                                            Name ="Label63"
                                            Caption ="Selected"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =22380
                                            LayoutCachedTop =3360
                                            LayoutCachedWidth =23280
                                            LayoutCachedHeight =3900
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =17254
                                    Top =3570
                                    Width =1822
                                    FontSize =9
                                    TabIndex =14
                                    BackColor =13434879
                                    BorderColor =12632256
                                    Name ="SearchSchedule"
                                    FontName ="Calibri"
                                    OnLostFocus ="[Event Procedure]"
                                    OnChange ="[Event Procedure]"

                                    LayoutCachedLeft =17254
                                    LayoutCachedTop =3570
                                    LayoutCachedWidth =19076
                                    LayoutCachedHeight =3810
                                End
                                Begin ToggleButton
                                    OverlapFlags =215
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =17254
                                    Top =3810
                                    Width =1822
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =15
                                    ForeColor =8210719
                                    Name ="ToggleSortSchedule"
                                    DefaultValue ="0"
                                    Caption ="Schedule"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =17254
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =19076
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin ToggleButton
                                    OverlapFlags =247
                                    TextFontCharSet =177
                                    TextFontFamily =0
                                    Left =15432
                                    Top =3810
                                    Width =300
                                    Height =240
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =16
                                    ForeColor =8210719
                                    Name ="ToggleSortuStatus"
                                    DefaultValue ="0"
                                    Caption ="Sts"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =15432
                                    LayoutCachedTop =3810
                                    LayoutCachedWidth =15732
                                    LayoutCachedHeight =4050
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    Shape =0
                                    Bevel =0
                                    BackColor =15527148
                                    BackThemeColorIndex =-1
                                    BackTint =100.0
                                    OldBorderStyle =1
                                    BorderWidth =1
                                    BorderColor =10855845
                                    BorderThemeColorIndex =-1
                                    BorderTint =100.0
                                    ThemeFontIndex =-1
                                    HoverColor =13017476
                                    HoverThemeColorIndex =-1
                                    HoverTint =100.0
                                    PressedColor =10319446
                                    PressedThemeColorIndex =-1
                                    PressedShade =100.0
                                    HoverForeColor =8210719
                                    HoverForeThemeColorIndex =-1
                                    HoverForeTint =100.0
                                    PressedForeThemeColorIndex =-1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =19560
                                    Top =9960
                                    Width =2527
                                    Height =420
                                    FontSize =11
                                    TabIndex =17
                                    Name ="MaketitleRequest"
                                    Caption ="Make Title Request"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
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

                                    LayoutCachedLeft =19560
                                    LayoutCachedTop =9960
                                    LayoutCachedWidth =22087
                                    LayoutCachedHeight =10380
                                    PictureCaptionArrangement =5
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =-2147483633
                                    BorderThemeColorIndex =1
                                    HoverThemeColorIndex =4
                                    HoverTint =80.0
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeThemeColorIndex =2
                                    PressedForeThemeColorIndex =2
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-625
                                    Overlaps =1
                                End
                                Begin OptionGroup
                                    OverlapFlags =215
                                    Left =19380
                                    Top =4020
                                    Width =2934
                                    Height =3435
                                    TabIndex =18
                                    Name ="RequestType"
                                    DefaultValue ="1"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =19380
                                    LayoutCachedTop =4020
                                    LayoutCachedWidth =22314
                                    LayoutCachedHeight =7455
                                    Begin
                                        Begin Label
                                            SpecialEffect =3
                                            BackStyle =1
                                            OldBorderStyle =1
                                            OverlapFlags =215
                                            TextAlign =2
                                            Left =19500
                                            Top =3900
                                            Width =1785
                                            Height =240
                                            Name ="Label588"
                                            Caption ="Select Reason"
                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =3900
                                            LayoutCachedWidth =21285
                                            LayoutCachedHeight =4140
                                            BorderThemeColorIndex =0
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =4260
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            OptionValue =1
                                            Name ="Toggle590"
                                            Caption ="1: Lease Terminated"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =4260
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =4635
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =4688
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =1
                                            OptionValue =2
                                            Name ="Toggle591"
                                            Caption ="2: Retitle the Vehicle"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =4688
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =5063
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =5116
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =2
                                            OptionValue =3
                                            Name ="Toggle592"
                                            Caption ="3: Purchased by Owner"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =5116
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =5491
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =5544
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =3
                                            OptionValue =4
                                            Name ="Toggle593"
                                            Caption ="4: Re-Registration"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =5544
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =5919
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =5972
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =4
                                            OptionValue =5
                                            Name ="Toggle594"
                                            Caption ="5: OK Affidavit of Lien Relse"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =5972
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =6347
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =223
                                            TextFontCharSet =177
                                            TextFontFamily =0
                                            Left =19500
                                            Top =6400
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =5
                                            OptionValue =6
                                            Name ="Toggle606"
                                            Caption ="6 Affidavit of Lien Release"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =6400
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =6775
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =215
                                            Left =19500
                                            Top =6828
                                            Width =2547
                                            Height =375
                                            FontSize =11
                                            FontWeight =400
                                            TabIndex =6
                                            OptionValue =7
                                            Name ="Toggle609"
                                            Caption ="7 Other"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120

                                            LayoutCachedLeft =19500
                                            LayoutCachedTop =6828
                                            LayoutCachedWidth =22047
                                            LayoutCachedHeight =7203
                                            ForeThemeColorIndex =2
                                            ForeTint =100.0
                                            Shape =0
                                            Bevel =0
                                            BackColor =-2147483633
                                            BackThemeColorIndex =-1
                                            BackTint =100.0
                                            OldBorderStyle =1
                                            BorderThemeColorIndex =1
                                            BorderTint =100.0
                                            ThemeFontIndex =-1
                                            HoverColor =13277810
                                            HoverTint =80.0
                                            PressedColor =10250042
                                            PressedShade =80.0
                                            HoverForeThemeColorIndex =1
                                            HoverForeTint =100.0
                                            Shadow =-1
                                            QuickStyle =30
                                            QuickStyleMask =-625
                                            Overlaps =1
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =19320
                                    Top =7980
                                    Width =4140
                                    Height =1530
                                    FontSize =10
                                    TabIndex =19
                                    BackColor =13434879
                                    Name ="OtherReason"
                                    FontName ="Calibri"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    TextFormat =1

                                    LayoutCachedLeft =19320
                                    LayoutCachedTop =7980
                                    LayoutCachedWidth =23460
                                    LayoutCachedHeight =9510
                                    RowStart =1
                                    RowEnd =1
                                    ColumnEnd =5
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    ThemeFontIndex =1
                                    ForeThemeColorIndex =0
                                    ForeTint =75.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =19320
                                            Top =7740
                                            Width =1410
                                            Height =225
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label6"
                                            Caption ="Other Reason"
                                            FontName ="Calibri"
                                            LeftPadding =0
                                            TopPadding =0
                                            RightPadding =0
                                            BottomPadding =0
                                            LayoutCachedLeft =19320
                                            LayoutCachedTop =7740
                                            LayoutCachedWidth =20730
                                            LayoutCachedHeight =7965
                                            ThemeFontIndex =1
                                            BorderThemeColorIndex =0
                                            BorderTint =50.0
                                            ForeThemeColorIndex =0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =12180
                                    Top =2700
                                    Width =3090
                                    Height =360
                                    FontSize =12
                                    FontWeight =700
                                    Name ="Label81"
                                    Caption ="Make a Title Request Form"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =12180
                                    LayoutCachedTop =2700
                                    LayoutCachedWidth =15270
                                    LayoutCachedHeight =3060
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =0
                                    BorderTint =50.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2459
                                    Top =8880
                                    Width =9120
                                    Height =3660
                                    TabIndex =20
                                    Name ="BankWireDetails_Sub"
                                    SourceObject ="Form.BankWireDetails_Sub"
                                    LinkChildFields ="BankID"
                                    LinkMasterFields ="BankID"

                                    LayoutCachedLeft =2459
                                    LayoutCachedTop =8880
                                    LayoutCachedWidth =11579
                                    LayoutCachedHeight =12540
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2460
                            Top =2475
                            Width =21045
                            Height =10125
                            BorderColor =0
                            Name ="Locations"
                            GridlineColor =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12600
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2745
                                    Top =2535
                                    Width =19980
                                    Height =6225
                                    Name ="Bank_Locations"
                                    SourceObject ="Form.Bank_Locations"
                                    LinkChildFields ="ClientREF"
                                    LinkMasterFields ="BankID"

                                    LayoutCachedLeft =2745
                                    LayoutCachedTop =2535
                                    LayoutCachedWidth =22725
                                    LayoutCachedHeight =8760
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2460
                            Top =2475
                            Width =21045
                            Height =10125
                            BorderColor =0
                            Name ="Notes tab"
                            EventProcPrefix ="Notes_tab"
                            Caption ="Notes"
                            GridlineColor =0
                            ImageData = Begin
                                0x00000000
                            End
                            LayoutCachedLeft =2460
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12600
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2460
                            Top =2475
                            Width =21045
                            Height =10125
                            Name ="RAL Conditions"
                            EventProcPrefix ="RAL_Conditions"
                            LayoutCachedLeft =2460
                            LayoutCachedTop =2475
                            LayoutCachedWidth =23505
                            LayoutCachedHeight =12600
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2460
                                    Top =2580
                                    Width =20640
                                    Height =10020
                                    Name ="RALConditions"
                                    SourceObject ="Form.RALConditions"
                                    LinkChildFields ="BankID"
                                    LinkMasterFields ="BankID"

                                    LayoutCachedLeft =2460
                                    LayoutCachedTop =2580
                                    LayoutCachedWidth =23100
                                    LayoutCachedHeight =12600
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Top =270
                    Width =3870
                    Height =270
                    FontSize =10
                    Name ="BankName"
                    ControlSource ="BankName"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =3495
                    LayoutCachedTop =270
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2490
                            Top =270
                            Width =945
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label138"
                            Caption ="Name:"
                            FontName ="Calibri"
                            GroupTable =3
                            LayoutCachedLeft =2490
                            LayoutCachedTop =270
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =540
                            LayoutGroup =1
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Top =600
                    Width =3870
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="clientDBA"
                    ControlSource ="BankDBA"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =3495
                    LayoutCachedTop =600
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =870
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2490
                            Top =600
                            Width =945
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label139"
                            Caption ="d/b/a:"
                            FontName ="Calibri"
                            GroupTable =3
                            LayoutCachedLeft =2490
                            LayoutCachedTop =600
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =870
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Top =930
                    Width =3870
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="Address1"
                    ControlSource ="Address1"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =3495
                    LayoutCachedTop =930
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =1200
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2490
                            Top =930
                            Width =945
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label140"
                            Caption ="Address1:"
                            FontName ="Calibri"
                            GroupTable =3
                            LayoutCachedLeft =2490
                            LayoutCachedTop =930
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =1200
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Top =1260
                    Width =3870
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="Address2"
                    ControlSource ="Address2"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =3495
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =1530
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2490
                            Top =1260
                            Width =945
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label141"
                            Caption ="Address2:"
                            FontName ="Calibri"
                            GroupTable =3
                            LayoutCachedLeft =2490
                            LayoutCachedTop =1260
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =1530
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3495
                    Top =1590
                    Width =2010
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="City"
                    ControlSource ="City"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =3495
                    LayoutCachedTop =1590
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2490
                            Top =1590
                            Width =945
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label142"
                            Caption ="City/St/Zip:"
                            FontName ="Calibri"
                            GroupTable =3
                            LayoutCachedLeft =2490
                            LayoutCachedTop =1590
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =1860
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =4
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5565
                    Top =1590
                    Width =690
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="State"
                    ControlSource ="State"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_States"
                    StatusBarText ="lookup states table"
                    FontName ="Calibri"
                    GroupTable =3

                    LayoutCachedLeft =5565
                    LayoutCachedTop =1590
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6315
                    Top =1590
                    Width =1050
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="Zip"
                    ControlSource ="Zip"
                    FontName ="Calibri"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =3

                    LayoutCachedLeft =6315
                    LayoutCachedTop =1590
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =1860
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8730
                    Top =930
                    Width =1605
                    Height =270
                    FontSize =9
                    TabIndex =10
                    BorderColor =11830108
                    Name ="Phone"
                    ControlSource ="Phone"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =10

                    LayoutCachedLeft =8730
                    LayoutCachedTop =930
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =1200
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =7605
                            Top =930
                            Width =1065
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label27"
                            Caption ="Main Phone:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =7605
                            LayoutCachedTop =930
                            LayoutCachedWidth =8670
                            LayoutCachedHeight =1200
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11040
                    Top =930
                    Width =1500
                    Height =270
                    FontSize =9
                    TabIndex =11
                    BorderColor =11830108
                    Name ="Fax"
                    ControlSource ="Fax"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =10
                    RightPadding =45

                    LayoutCachedLeft =11040
                    LayoutCachedTop =930
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =1200
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10395
                            Top =930
                            Width =585
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label31"
                            Caption ="Fax:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =10395
                            LayoutCachedTop =930
                            LayoutCachedWidth =10980
                            LayoutCachedHeight =1200
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8730
                    Top =600
                    Width =3810
                    Height =270
                    FontSize =9
                    TabIndex =9
                    BorderColor =11830108
                    ForeColor =1279872587
                    Name ="Website"
                    ControlSource ="Website"
                    FontName ="Calibri"
                    GroupTable =10
                    RightPadding =45

                    LayoutCachedLeft =8730
                    LayoutCachedTop =600
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =870
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =7605
                            Top =600
                            Width =1065
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label35"
                            Caption ="Website:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =7605
                            LayoutCachedTop =600
                            LayoutCachedWidth =8670
                            LayoutCachedHeight =870
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11655
                    Top =1260
                    Width =885
                    Height =270
                    FontSize =9
                    TabIndex =13
                    BorderColor =11830108
                    Name ="BecameBank"
                    ControlSource ="BecameBank"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GroupTable =10
                    RightPadding =45

                    LayoutCachedLeft =11655
                    LayoutCachedTop =1260
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =1530
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =10395
                            Top =1260
                            Width =1192
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label43"
                            Caption ="Became Bank:"
                            FontName ="Calibri"
                            GroupTable =10
                            RightPadding =45
                            LayoutCachedLeft =10395
                            LayoutCachedTop =1260
                            LayoutCachedWidth =11587
                            LayoutCachedHeight =1530
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =4
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9360
                    Top =1260
                    Width =975
                    Height =270
                    FontSize =9
                    TabIndex =12
                    BorderColor =11830108
                    Name ="LeaseSTofCorp"
                    ControlSource ="CorpState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY lookup_States.State"
                        "Abbrev;"
                    FontName ="Calibri"
                    GroupTable =10

                    LayoutCachedLeft =9360
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =1530
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =7605
                            Top =1260
                            Width =1695
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label249"
                            Caption ="State Incorporated in:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =7605
                            LayoutCachedTop =1260
                            LayoutCachedWidth =9300
                            LayoutCachedHeight =1530
                            RowStart =3
                            RowEnd =3
                            ColumnEnd =1
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11040
                    Top =270
                    Width =1500
                    Height =270
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =8872517
                    Name ="TaxID"
                    ControlSource ="TaxID"
                    FontName ="Calibri"
                    GroupTable =10
                    RightPadding =45

                    LayoutCachedLeft =11040
                    LayoutCachedTop =270
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10395
                            Top =270
                            Width =585
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label248"
                            Caption ="Tax ID:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =10395
                            LayoutCachedTop =270
                            LayoutCachedWidth =10980
                            LayoutCachedHeight =540
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8730
                    Top =270
                    Width =1605
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BackColor =15590879
                    BorderColor =11830108
                    Name ="BankShortName"
                    ControlSource ="BankShortName"
                    StatusBarText ="Short Name for Reporting"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Short Name is required"
                    FontName ="Calibri"
                    GroupTable =10

                    LayoutCachedLeft =8730
                    LayoutCachedTop =270
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =7605
                            Top =270
                            Width =1065
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label358"
                            Caption ="Short Name:"
                            FontName ="Calibri"
                            GroupTable =10
                            LayoutCachedLeft =7605
                            LayoutCachedTop =270
                            LayoutCachedWidth =8670
                            LayoutCachedHeight =540
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19050
                    Top =1665
                    Width =907
                    FontSize =9
                    FontWeight =700
                    TabIndex =22
                    BorderColor =11830108
                    Name ="clientType"
                    ControlSource ="BankType"
                    ValidationRule ="Is Not Null"
                    ValidationText ="CLIENT TYPE REQUIRED"
                    DefaultValue ="\"BANK\""
                    FontName ="Calibri"
                    BottomPadding =45

                    LayoutCachedLeft =19050
                    LayoutCachedTop =1665
                    LayoutCachedWidth =19957
                    LayoutCachedHeight =1905
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =18420
                            Top =1665
                            Width =570
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label37"
                            Caption ="Type:"
                            FontName ="Calibri"
                            BottomPadding =45
                            LayoutCachedLeft =18420
                            LayoutCachedTop =1665
                            LayoutCachedWidth =18990
                            LayoutCachedHeight =1905
                            RowStart =4
                            RowEnd =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =12720
                    Top =180
                    Width =3615
                    Height =1740
                    Name ="Box377"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =180
                    LayoutCachedWidth =16335
                    LayoutCachedHeight =1920
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =16710
                    Top =510
                    Width =3810
                    FontSize =9
                    TabIndex =16
                    BorderColor =11830108
                    Name ="AssigneeBank"
                    ControlSource ="BankType"
                    RowSourceType ="Value List"
                    RowSource ="1;Assignee;3;Interim;2;Assignee-Interim;4;Non-Bank;5;Competitor;6;Other-Delete"
                    ColumnWidths ="0;3600"
                    StatusBarText ="to separate competitors from Assignees"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Financial Institution Type is required"
                    DefaultValue ="1"
                    FontName ="Calibri"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16710
                    LayoutCachedTop =510
                    LayoutCachedWidth =20520
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =2
                            Left =16605
                            Top =240
                            Width =3975
                            Height =585
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483607
                            Name ="Label408"
                            Caption ="Financial Institution Type:"
                            FontName ="Calibri"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =16605
                            LayoutCachedTop =240
                            LayoutCachedWidth =20580
                            LayoutCachedHeight =825
                            BackThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =450
                    Width =2160
                    Height =12150
                    FontSize =9
                    TabIndex =15
                    BackColor =16249583
                    Name ="BankList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName AS Bank FROM Banks ORDER BY Banks.BankS"
                        "hortName;"
                    ColumnWidths ="0;1584"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =450
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =12600
                    BorderThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =75
                    Top =15
                    Width =2160
                    Height =405
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label447"
                    Caption ="Banks "
                    FontName ="Calibri"
                    LayoutCachedLeft =75
                    LayoutCachedTop =15
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16665
                    Top =1230
                    Width =1860
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    Name ="Text448"
                    ControlSource ="PortfolioNumber"
                    StatusBarText ="Short Name for Reporting"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Short Name is required"
                    FontName ="Calibri"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16665
                    LayoutCachedTop =1230
                    LayoutCachedWidth =18525
                    LayoutCachedHeight =1470
                    BorderThemeColorIndex =7
                    BorderShade =50.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =2
                            Left =16605
                            Top =960
                            Width =1950
                            Height =585
                            FontSize =9
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label449"
                            Caption ="Portfolio SYND:"
                            FontName ="Calibri"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =16605
                            LayoutCachedTop =960
                            LayoutCachedWidth =18555
                            LayoutCachedHeight =1545
                            BackThemeColorIndex =8
                            BackShade =50.0
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =7560
                    Top =180
                    Width =5115
                    Height =1740
                    Name ="Box450"
                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =1920
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =2445
                    Top =180
                    Width =5040
                    Height =1740
                    Name ="Box451"
                    LayoutCachedLeft =2445
                    LayoutCachedTop =180
                    LayoutCachedWidth =7485
                    LayoutCachedHeight =1920
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =16395
                    Top =180
                    Width =4395
                    Height =1740
                    Name ="Box452"
                    LayoutCachedLeft =16395
                    LayoutCachedTop =180
                    LayoutCachedWidth =20790
                    LayoutCachedHeight =1920
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2475
                    Top =45
                    Width =1290
                    Height =210
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label120"
                    Caption ="Bank Information"
                    FontName ="Calibri"
                    LayoutCachedLeft =2475
                    LayoutCachedTop =45
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20070
                    Top =1650
                    Width =690
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =21
                    BackColor =15707392
                    ForeColor =5026082
                    Name ="BankID"
                    ControlSource ="BankID"
                    FontName ="Calibri"

                    LayoutCachedLeft =20070
                    LayoutCachedTop =1650
                    LayoutCachedWidth =20760
                    LayoutCachedHeight =1890
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18705
                    Top =1230
                    Width =1860
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    Name ="Text558"
                    ControlSource ="PortfolioNumberInterm"
                    StatusBarText ="Short Name for Reporting"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Short Name is required"
                    FontName ="Calibri"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18705
                    LayoutCachedTop =1230
                    LayoutCachedWidth =20565
                    LayoutCachedHeight =1470
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =18645
                            Top =960
                            Width =1950
                            Height =585
                            FontSize =9
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label559"
                            Caption ="Portfolio Interim:"
                            FontName ="Calibri"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18645
                            LayoutCachedTop =960
                            LayoutCachedWidth =20595
                            LayoutCachedHeight =1545
                            BackThemeColorIndex =6
                            BackShade =75.0
                            BorderThemeColorIndex =6
                            BorderShade =75.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =8070
                    Top =1605
                    Width =1455
                    Height =240
                    FontSize =10
                    TabIndex =19
                    Name ="DrawDownMaker"
                    Caption ="Draw Maker"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open Draw Down Maker Form"
                    BottomPadding =45

                    LayoutCachedLeft =8070
                    LayoutCachedTop =1605
                    LayoutCachedWidth =9525
                    LayoutCachedHeight =1845
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10710
                    Top =1590
                    Width =1455
                    Height =240
                    FontSize =10
                    TabIndex =20
                    Name ="DrawDownManager"
                    Caption ="Draw Manager"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open Draw Down Maker Form"
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =10710
                    LayoutCachedTop =1590
                    LayoutCachedWidth =12165
                    LayoutCachedHeight =1830
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12735
                    Top =285
                    Width =3555
                    Height =1293
                    TabIndex =14
                    BackColor =15858167
                    Name ="RemitAddress"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12735
                    LayoutCachedTop =285
                    LayoutCachedWidth =16290
                    LayoutCachedHeight =1578
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =12795
                    Top =30
                    Width =1170
                    Height =210
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label580"
                    Caption ="Remit Address"
                    FontName ="Calibri"
                    LayoutCachedLeft =12795
                    LayoutCachedTop =30
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =240
                End
                Begin EmptyCell
                    Left =1575
                    Width =0
                    Height =360
                    Name ="EmptyCell595"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1575
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =360
                    LayoutGroup =4
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =1575
                    Width =0
                    Height =360
                    Name ="EmptyCell596"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1575
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =1575
                    Width =0
                    Height =360
                    Name ="EmptyCell597"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1575
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =360
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =1575
                    Width =5745
                    Height =360
                    Name ="EmptyCell598"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1575
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =360
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =12
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
' See "Bank_View.cls"
