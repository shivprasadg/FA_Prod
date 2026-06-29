Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =10
    ItemSuffix =401
    Left =12915
    Top =5730
    Right =27570
    Bottom =10890
    Filter ="[dealID]=1491"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x9bad32bfd43ce440
    End
    RecordSource ="DealsQuery"
    Caption ="View/Edit Deals"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnGotFocus ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Deals"
    ShowPageMargins =0
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1035
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Width =7665
                    Height =480
                    FontSize =18
                    Name ="companyAndDBA"
                    ControlSource ="CompanyAndDBA"

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =8925
                    LayoutCachedHeight =480
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1275
                    Top =510
                    Width =7665
                    Height =255
                    ColumnWidth =2670
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =12632256
                    Name ="dealShortDesc"
                    ControlSource ="dealShortDesc"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =510
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =765
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9420
                    Top =240
                    Width =1980
                    Height =228
                    TabIndex =1
                    Name ="ClientViewFormBtn"
                    Caption ="Open Client's Form"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9420
                    LayoutCachedTop =240
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =468
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
                    OverlapFlags =93
                    TextAlign =3
                    Left =120
                    Width =1020
                    Height =525
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Deals"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =525
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Top =405
                    Width =1125
                    Height =240
                    Name ="Label227"
                    Caption ="Dashboard"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =405
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =645
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =13080
                    Top =240
                    Width =555
                    Height =228
                    TabIndex =2
                    Name ="GoHome"
                    Caption ="Home"
                    FontName ="Calibri"
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

                    LayoutCachedLeft =13080
                    LayoutCachedTop =240
                    LayoutCachedWidth =13635
                    LayoutCachedHeight =468
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
                    Left =13647
                    Top =240
                    Width =555
                    Height =228
                    TabIndex =3
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =13647
                    LayoutCachedTop =240
                    LayoutCachedWidth =14202
                    LayoutCachedHeight =468
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
            Height =4140
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11040
                    Top =3540
                    Width =570
                    Height =255
                    FontSize =7
                    FontWeight =700
                    TabIndex =24
                    Name ="clientRef"
                    ControlSource ="clientRef"
                    FontName ="Calibri"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11040
                    LayoutCachedTop =3540
                    LayoutCachedWidth =11610
                    LayoutCachedHeight =3795
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10440
                            Top =3540
                            Width =600
                            Height =255
                            FontSize =7
                            FontWeight =700
                            Name ="Label166"
                            Caption ="CGroup"
                            FontName ="Calibri"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10440
                            LayoutCachedTop =3540
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =3795
                            BorderThemeColorIndex =1
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12195
                    Top =3540
                    Width =570
                    Height =255
                    FontSize =7
                    FontWeight =700
                    TabIndex =25
                    Name ="dealIDd"
                    ControlSource ="dealID"
                    FontName ="Calibri"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12195
                    LayoutCachedTop =3540
                    LayoutCachedWidth =12765
                    LayoutCachedHeight =3795
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =11625
                            Top =3540
                            Width =570
                            Height =255
                            FontSize =7
                            FontWeight =700
                            Name ="Label105"
                            Caption ="DealID"
                            FontName ="Calibri"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =11625
                            LayoutCachedTop =3540
                            LayoutCachedWidth =12195
                            LayoutCachedHeight =3795
                            BorderThemeColorIndex =1
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            ForeShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =8595
                    Width =5550
                    Height =3195
                    BackColor =11830108
                    Name ="Box124"
                    LayoutCachedLeft =8595
                    LayoutCachedWidth =14145
                    LayoutCachedHeight =3195
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =900
                    Top =810
                    Width =1020
                    Height =210
                    Name ="FAID"
                    ControlSource ="FAID"

                    LayoutCachedLeft =900
                    LayoutCachedTop =810
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =1020
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =900
                            Top =600
                            Width =1020
                            Height =210
                            Name ="Label54"
                            Caption ="FAID"
                            LayoutCachedLeft =900
                            LayoutCachedTop =600
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =810
                            ForeThemeColorIndex =4
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =810
                    Width =1365
                    Height =210
                    TabIndex =1
                    Name ="Transtype"
                    ControlSource ="transType"
                    RowSourceType ="Value List"
                    StatusBarText ="link to lookup_transTypes table"
                    InheritValueList =1

                    LayoutCachedLeft =1965
                    LayoutCachedTop =810
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =1020
                    DatasheetCaption ="Transtype"
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1965
                            Top =600
                            Width =1380
                            Height =210
                            Name ="Label174"
                            Caption ="Type"
                            LayoutCachedLeft =1965
                            LayoutCachedTop =600
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =810
                            ForeThemeColorIndex =4
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =900
                    Top =1512
                    Width =900
                    FontWeight =700
                    TabIndex =4
                    Name ="DealLOEx"
                    ControlSource ="DealLOEx"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Salesperson.SPInitials FROM Salesperson WHERE (((Salesperson.Type)=\"1\")"
                        " AND ((Salesperson.Inactive)=False)) ORDER BY Salesperson.SPInitials;"
                    StatusBarText ="LO Exec at time of Proposal"
                    ValidationRule ="Is Not Null"
                    ValidationText ="LO Exec is required to be entered on deal"
                    GroupTable =4

                    LayoutCachedLeft =900
                    LayoutCachedTop =1512
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =1752
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =900
                    Top =2172
                    Width =1152
                    Height =252
                    TabIndex =7
                    Name ="RFPDate"
                    ControlSource ="dealRFPDate"
                    Format ="Short Date"
                    GroupTable =11

                    LayoutCachedLeft =900
                    LayoutCachedTop =2172
                    LayoutCachedWidth =2052
                    LayoutCachedHeight =2424
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2124
                    Top =2172
                    Width =1152
                    Height =252
                    TabIndex =8
                    Name ="dealLPAFDue"
                    ControlSource ="dealLPAFDue"
                    Format ="Short Date"
                    GroupTable =11

                    LayoutCachedLeft =2124
                    LayoutCachedTop =2172
                    LayoutCachedWidth =3276
                    LayoutCachedHeight =2424
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3348
                    Top =2172
                    Width =1152
                    Height =252
                    TabIndex =9
                    Name ="dealLPAFcomp"
                    ControlSource ="dealLPAFcomp"
                    Format ="Short Date"
                    StatusBarText ="2/29/12 added date completed"
                    GroupTable =11

                    LayoutCachedLeft =3348
                    LayoutCachedTop =2172
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =2424
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4572
                    Top =2172
                    Width =1152
                    Height =252
                    TabIndex =10
                    Name ="dealLPAFSent"
                    ControlSource ="dealLPAFSent"
                    Format ="Short Date"
                    GroupTable =11
                    OnDirtyEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Condition ="[Forms]![deals_View]![transtype]<>\"proposal\""
                            Action ="MsgBox"
                            Argument ="Transaction must be Proposal to enter the LPAF Sent date"
                            Argument ="-1"
                            Argument ="1"
                            Argument ="Transaction Type Error"
                        End
                        Begin
                            Condition ="..."
                            Action ="GoToControl"
                            Argument ="[transtype]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"dealLPAFSent\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition"
                                ">[Forms]![deals_View]![transt"
                        End
                        Begin
                            Comment ="_AXL:ype]&lt;&gt;\"proposal\"</Condition><Statements><Action Name=\"MessageBox\""
                                "><Argument Name=\"Message\">Transaction must be Proposal to enter the LPAF Sent "
                                "date</Argument><Argument Name=\"Type\">Critical</Argument><Argument Name=\"Title"
                                "\">Transaction Type Err"
                        End
                        Begin
                            Comment ="_AXL:or</Argument></Action><Action Name=\"GoToControl\"><Argument Name=\"Control"
                                "Name\">[transtype]</Argument></Action></Statements></If></ConditionalBlock></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4572
                    LayoutCachedTop =2172
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =2424
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5796
                    Top =2172
                    Width =1152
                    Height =252
                    TabIndex =11
                    BackColor =14151142
                    Name ="dealEstDelDate"
                    ControlSource ="dealEstDelDate"
                    Format ="Short Date"
                    GroupTable =11

                    LayoutCachedLeft =5796
                    LayoutCachedTop =2172
                    LayoutCachedWidth =6948
                    LayoutCachedHeight =2424
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =900
                    Top =2832
                    Width =780
                    TabIndex =17
                    Name ="dealMPY"
                    ControlSource ="dealMPY"
                    Format ="Standard"
                    GroupTable =32

                    LayoutCachedLeft =900
                    LayoutCachedTop =2832
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1752
                    Top =2832
                    Width =480
                    TabIndex =18
                    Name ="dealTerm"
                    ControlSource ="dealTerm"
                    GroupTable =32

                    LayoutCachedLeft =1752
                    LayoutCachedTop =2832
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2304
                    Top =2832
                    Width =1020
                    FontWeight =700
                    TabIndex =19
                    Name ="dealUnitQty"
                    ControlSource ="dealUnitQty"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =32

                    LayoutCachedLeft =2304
                    LayoutCachedTop =2832
                    LayoutCachedWidth =3324
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3396
                    Top =2832
                    TabIndex =20
                    Name ="dealUnitPrice"
                    ControlSource ="dealUnitPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =32

                    LayoutCachedLeft =3396
                    LayoutCachedTop =2832
                    LayoutCachedWidth =4836
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2280
                    Top =3300
                    Width =2520
                    Height =255
                    TabIndex =26
                    Name ="dealSchedNum"
                    ControlSource ="dealSchedNum"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3300
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =3555
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =900
                            Top =3315
                            Width =1380
                            Height =240
                            Name ="Label133"
                            Caption ="Atlaas Schedule(s)"
                            LayoutCachedLeft =900
                            LayoutCachedTop =3315
                            LayoutCachedWidth =2280
                            LayoutCachedHeight =3555
                            ForeThemeColorIndex =4
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10140
                    Top =90
                    Width =1125
                    Height =255
                    TabIndex =2
                    BorderColor =13017476
                    Name ="dealOutcomeDate"
                    ControlSource ="dealOutcomeDate"
                    Format ="Short Date"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =90
                    LayoutCachedWidth =11265
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            Left =8640
                            Top =90
                            Width =1395
                            Height =210
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label121"
                            Caption ="Outcome Date"
                            LayoutCachedLeft =8640
                            LayoutCachedTop =90
                            LayoutCachedWidth =10035
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10770
                    Top =795
                    Width =2805
                    Height =255
                    TabIndex =6
                    BorderColor =13017476
                    Name ="LostTo"
                    ControlSource ="LostTo"
                    RowSourceType ="Value List"
                    RowSource ="None;Ownership;T.R.A.C.;Finance Lease;FSL"
                    AllowValueListEdits =1

                    LayoutCachedLeft =10770
                    LayoutCachedTop =795
                    LayoutCachedWidth =13575
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =3
                            Left =9060
                            Top =780
                            Width =1665
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label188"
                            Caption ="Aquisition Method:"
                            LayoutCachedLeft =9060
                            LayoutCachedTop =780
                            LayoutCachedWidth =10725
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListRows =8
                    Left =10770
                    Top =1125
                    Width =2805
                    Height =255
                    TabIndex =12
                    BorderColor =13017476
                    Name ="Lost To:"
                    ControlSource ="Lost To:"
                    RowSourceType ="Value List"
                    RowSource ="BofA;GE;PNC;TCF;USBank;Wells;5/3;Unknown Bank;Ameriquest;LeasePlan;PHH;AIM;Pensk"
                        "e;Paccar;Ryder;Volvo;Ownership;Other"
                    ColumnWidths ="720"
                    StatusBarText ="actual entity"
                    EventProcPrefix ="Lost_To_"
                    Format =">"
                    AllowValueListEdits =1

                    LayoutCachedLeft =10770
                    LayoutCachedTop =1125
                    LayoutCachedWidth =13575
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =3
                            Left =9615
                            Top =1125
                            Width =1110
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label192"
                            Caption ="Competitor:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =1125
                            LayoutCachedWidth =10725
                            LayoutCachedHeight =1365
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10770
                    Top =1440
                    Width =915
                    Height =255
                    TabIndex =13
                    BorderColor =13017476
                    Name ="LostTerm"
                    ControlSource ="LostTerm"
                    StatusBarText ="competitor term"

                    LayoutCachedLeft =10770
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =1695
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =3
                            Left =10050
                            Top =1440
                            Width =675
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label189"
                            Caption ="Term: "
                            LayoutCachedLeft =10050
                            LayoutCachedTop =1440
                            LayoutCachedWidth =10725
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13275
                    Top =1440
                    Width =705
                    Height =255
                    TabIndex =14
                    BorderColor =13017476
                    Name ="LostResid"
                    ControlSource ="LostResid"
                    Format ="Percent"
                    StatusBarText ="if T.R.A.C. Lease, residual %"
                    ControlTipText ="if T.R.A.C. Lease, residual %"

                    LayoutCachedLeft =13275
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1695
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            Left =11775
                            Top =1440
                            Width =1275
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="TRACRESID"
                            Caption ="if T.R.A.C. to: "
                            ControlTipText ="if T.R.A.C. Lease, residual %"
                            LayoutCachedLeft =11775
                            LayoutCachedTop =1440
                            LayoutCachedWidth =13050
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10770
                    Top =1755
                    Width =915
                    Height =255
                    TabIndex =15
                    BorderColor =13017476
                    Name ="LostRate"
                    ControlSource ="LostRate"
                    Format ="Percent"
                    StatusBarText ="competitor rate"

                    LayoutCachedLeft =10770
                    LayoutCachedTop =1755
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =2010
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =3
                            Left =9690
                            Top =1755
                            Width =1035
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label193"
                            Caption ="Rate: "
                            LayoutCachedLeft =9690
                            LayoutCachedTop =1755
                            LayoutCachedWidth =10725
                            LayoutCachedHeight =1995
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =8700
                    Top =465
                    Width =5325
                    Height =225
                    FontWeight =700
                    Name ="Label187"
                    Caption ="Lost Deal Information"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =465
                    LayoutCachedWidth =14025
                    LayoutCachedHeight =690
                    ForeThemeColorIndex =1
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13260
                    Top =1785
                    Width =705
                    Height =255
                    TabIndex =16
                    BorderColor =13017476
                    Name ="LostImplicit Rt"
                    ControlSource ="LostImplicit Rt"
                    Format ="Percent"
                    EventProcPrefix ="LostImplicit_Rt"

                    LayoutCachedLeft =13260
                    LayoutCachedTop =1785
                    LayoutCachedWidth =13965
                    LayoutCachedHeight =2040
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            Left =11760
                            Top =1785
                            Width =1275
                            Height =240
                            FontWeight =700
                            ForeColor =16777215
                            Name ="ImpRT"
                            Caption ="Implicit Rate: "
                            LayoutCachedLeft =11760
                            LayoutCachedTop =1785
                            LayoutCachedWidth =13035
                            LayoutCachedHeight =2025
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9525
                    Top =2145
                    Width =4590
                    Height =1011
                    TabIndex =22
                    BorderColor =13017476
                    Name ="OutcomeReason"
                    ControlSource ="dealOutcomeReason"

                    LayoutCachedLeft =9525
                    LayoutCachedTop =2145
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =3156
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            Left =8685
                            Top =2130
                            Width =765
                            Height =210
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label155"
                            Caption ="Reason: "
                            LayoutCachedLeft =8685
                            LayoutCachedTop =2130
                            LayoutCachedWidth =9450
                            LayoutCachedHeight =2340
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =2820
                    Width =1620
                    Height =255
                    TabIndex =23
                    BorderColor =13017476
                    Name ="dealGrandTotal"
                    ControlSource ="dealGrandTotal"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="this is a calculated field"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =3075
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4908
                    Top =2832
                    Width =1560
                    TabIndex =21
                    Name ="Text182"
                    ControlSource ="=[dealunitqty]*[dealunitprice]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =32

                    LayoutCachedLeft =4908
                    LayoutCachedTop =2832
                    LayoutCachedWidth =6468
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =12390
                    Top =90
                    Height =255
                    TabIndex =3
                    BorderColor =13017476
                    Name ="dealOutcome"
                    ControlSource ="dealOutcome"
                    Format ="Short Date"
                    AfterUpdateEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Condition ="[Forms]![deals_View]![transtype]=\"opportunity\" And [Forms]![deals_View]![dealo"
                                "utcome]=\"Lost\""
                            Action ="MsgBox"
                            Argument ="Transaction must be a Proposal to enter Lost ; if Opportunity enter dead."
                            Argument ="-1"
                            Argument ="1"
                            Argument ="Transaction Type Error"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"dealOutcome\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition>"
                                "[Forms]![deals_View]![transty"
                        End
                        Begin
                            Comment ="_AXL:pe]=\"opportunity\" And [Forms]![deals_View]![dealoutcome]=\"Lost\"</Condit"
                                "ion><Statements><Action Name=\"MessageBox\"><Argument Name=\"Message\">Transacti"
                                "on must be a Proposal to enter Lost ; if Opportunity enter dead.</Argument><Argu"
                                "ment Name=\"Type\">Crit"
                        End
                        Begin
                            Comment ="_AXL:ical</Argument><Argument Name=\"Title\">Transaction Type Error</Argument></"
                                "Action></Statements></If></ConditionalBlock></Statements></UserInterfaceMacro>"
                        End
                    End
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Condition ="[Forms]![deals_View]![transtype]=\"opportunity\" And [Forms]![deals_View]![dealo"
                                "utcome]=\"Lost\""
                            Action ="MsgBox"
                            Argument ="Transaction must be a Proposal to enter Lost ; if Opportunity enter dead."
                            Argument ="-1"
                            Argument ="1"
                            Argument ="Transaction Type Error"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"dealOutcome\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition>"
                                "[Forms]![deals_View]![transty"
                        End
                        Begin
                            Comment ="_AXL:pe]=\"opportunity\" And [Forms]![deals_View]![dealoutcome]=\"Lost\"</Condit"
                                "ion><Statements><Action Name=\"MessageBox\"><Argument Name=\"Message\">Transacti"
                                "on must be a Proposal to enter Lost ; if Opportunity enter dead.</Argument><Argu"
                                "ment Name=\"Type\">Crit"
                        End
                        Begin
                            Comment ="_AXL:ical</Argument><Argument Name=\"Title\">Transaction Type Error</Argument></"
                                "Action></Statements></If></ConditionalBlock></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =12390
                    LayoutCachedTop =90
                    LayoutCachedWidth =13830
                    LayoutCachedHeight =345
                    DatasheetCaption ="Outcome"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            Left =11430
                            Top =120
                            Width =915
                            Height =210
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label122"
                            Caption ="Outcome"
                            LayoutCachedLeft =11430
                            LayoutCachedTop =120
                            LayoutCachedWidth =12345
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1872
                    Top =1512
                    Width =6300
                    TabIndex =5
                    Name ="Text200"
                    ControlSource ="dealShortDesc"
                    GroupTable =4

                    LayoutCachedLeft =1872
                    LayoutCachedTop =1512
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =1752
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =900
                    Top =1200
                    Width =900
                    Height =240
                    Name ="Label194"
                    Caption ="BFEX"
                    GroupTable =4
                    LayoutCachedLeft =900
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =1440
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1872
                    Top =1200
                    Width =6300
                    Height =240
                    Name ="Label201"
                    Caption ="Desc:"
                    GroupTable =4
                    LayoutCachedLeft =1872
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8172
                    LayoutCachedHeight =1440
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2124
                    Top =1860
                    Width =1152
                    Height =240
                    Name ="Label112"
                    Caption ="LPAF Due"
                    GroupTable =11
                    LayoutCachedLeft =2124
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3276
                    LayoutCachedHeight =2100
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3348
                    Top =1860
                    Width =1152
                    Height =240
                    Name ="Label179"
                    Caption ="LPAF Complete"
                    GroupTable =11
                    LayoutCachedLeft =3348
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =2100
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4572
                    Top =1860
                    Width =1152
                    Height =240
                    Name ="Label115"
                    Caption ="LPAF Sent"
                    GroupTable =11
                    LayoutCachedLeft =4572
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5724
                    LayoutCachedHeight =2100
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =900
                    Top =1860
                    Width =1152
                    Height =240
                    Name ="Label111"
                    Caption ="RFP Date"
                    GroupTable =11
                    LayoutCachedLeft =900
                    LayoutCachedTop =1860
                    LayoutCachedWidth =2052
                    LayoutCachedHeight =2100
                    LayoutGroup =2
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5796
                    Top =1860
                    Width =1152
                    Height =240
                    Name ="Label160"
                    Caption ="Est Delivery"
                    GroupTable =11
                    LayoutCachedLeft =5796
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6948
                    LayoutCachedHeight =2100
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =900
                    Top =2520
                    Width =780
                    Height =240
                    Name ="Label171"
                    Caption ="MPY"
                    GroupTable =32
                    LayoutCachedLeft =900
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =2760
                    LayoutGroup =3
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1752
                    Top =2520
                    Width =480
                    Height =240
                    Name ="Label172"
                    Caption ="Term"
                    GroupTable =32
                    LayoutCachedLeft =1752
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =2760
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2304
                    Top =2520
                    Width =1020
                    Height =240
                    Name ="Label149"
                    Caption ="Quantity"
                    GroupTable =32
                    LayoutCachedLeft =2304
                    LayoutCachedTop =2520
                    LayoutCachedWidth =3324
                    LayoutCachedHeight =2760
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3396
                    Top =2520
                    Width =1440
                    Height =240
                    Name ="Label150"
                    Caption ="Unit Price"
                    GroupTable =32
                    LayoutCachedLeft =3396
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4836
                    LayoutCachedHeight =2760
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4908
                    Top =2520
                    Width =1560
                    Height =240
                    Name ="Label183"
                    Caption ="Total Deal"
                    GroupTable =32
                    LayoutCachedLeft =4908
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6468
                    LayoutCachedHeight =2760
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =32
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =3600
                    Height =360
                    FontSize =11
                    Name ="Label254"
                    Caption ="LPAF Impoted Figures"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =420
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
' See "Deals_view.cls"
