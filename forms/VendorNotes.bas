Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9240
    DatasheetFontHeight =10
    ItemSuffix =194
    Left =8970
    Top =14520
    Right =18795
    Bottom =20610
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x76a516005854e640
    End
    RecordSource ="SELECT VendorNotes.VendorNoteID, VendorNotes.vnVendorID, VendorNotes.vnDate, Ven"
        "dorNotes.vnMemo, VendorNotes.vnFAUser FROM VendorNotes;"
    Caption ="notes"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Notes"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
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
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3060
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1740
                    Top =2700
                    Width =480
                    ColumnWidth =1440
                    TabIndex =5
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="VendorNoteID"
                    ControlSource ="VendorNoteID"
                    FontName ="Calibri"
                    BottomPadding =150

                    LayoutCachedLeft =1740
                    LayoutCachedTop =2700
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2940
                    ColumnStart =8
                    ColumnEnd =8
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2865
                    Top =60
                    Width =1275
                    Height =270
                    ColumnWidth =1860
                    FontSize =11
                    BackColor =11271167
                    BorderColor =12632256
                    ForeColor =0
                    Name ="vnDate"
                    ControlSource ="vnDate"
                    Format ="Short Date"
                    FontName ="Calibri"
                    GroupTable =14
                    BottomPadding =150

                    LayoutCachedLeft =2865
                    LayoutCachedTop =60
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =330
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =14
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2280
                            Top =60
                            Width =525
                            Height =270
                            FontSize =9
                            FontWeight =700
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label43"
                            Caption ="Date"
                            FontName ="Calibri"
                            GroupTable =14
                            BottomPadding =150
                            LayoutCachedLeft =2280
                            LayoutCachedTop =60
                            LayoutCachedWidth =2805
                            LayoutCachedHeight =330
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =14
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2280
                    Top =480
                    Width =6840
                    Height =2460
                    ColumnWidth =3000
                    FontSize =11
                    TabIndex =4
                    BackColor =16777215
                    BorderColor =12632256
                    ForeColor =0
                    Name ="vnMemo"
                    ControlSource ="vnMemo"
                    FontName ="Calibri"
                    GroupTable =13

                    LayoutCachedLeft =2280
                    LayoutCachedTop =480
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =2940
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =13
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =4560
                    Top =60
                    Width =1080
                    Height =270
                    FontSize =11
                    TabIndex =1
                    BackColor =11271167
                    BorderColor =12632256
                    ForeColor =0
                    Name ="vnFAUser"
                    ControlSource ="vnFAUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.OBSOLETE)=No)) ORDER BY E"
                        "mps.EmpInt;"
                    ColumnWidths ="720;2160"
                    StatusBarText ="entered by salesman"
                    ValidationRule ="Is Not Null"
                    ValidationText ="\"Entered by\" is required"
                    FontName ="Calibri"
                    GroupTable =14
                    BottomPadding =150

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =330
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =14
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4200
                            Top =60
                            Width =300
                            Height =270
                            FontSize =9
                            FontWeight =700
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label31"
                            Caption ="By"
                            FontName ="Calibri"
                            GroupTable =14
                            BottomPadding =150
                            LayoutCachedLeft =4200
                            LayoutCachedTop =60
                            LayoutCachedWidth =4500
                            LayoutCachedHeight =330
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =14
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =5700
                    Top =60
                    Width =1680
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="AddNoteBtn"
                    Caption ="Add a Note"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =14
                    BottomPadding =150
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =330
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =48840
                    BorderThemeColorIndex =1
                    HoverColor =48840
                    PressedColor =48840
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =14
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7440
                    Top =60
                    Width =1680
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GroupTable =14
                    BottomPadding =150

                    LayoutCachedLeft =7440
                    LayoutCachedTop =60
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =330
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =1
                    HoverColor =48840
                    PressedColor =48840
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =14
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1920
                    Top =480
                    Width =285
                    Height =1140
                    Name ="Label1744"
                    Caption =" └─ Notes"
                    FontName ="Segoe UI"
                    GroupTable =11
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1920
                    LayoutCachedTop =480
                    LayoutCachedWidth =2205
                    LayoutCachedHeight =1620
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1740
                    Top =2460
                    Width =480
                    TabIndex =7
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="vnVendorID"
                    ControlSource ="vnVendorID"
                    FontName ="Calibri"
                    BottomPadding =150

                    LayoutCachedLeft =1740
                    LayoutCachedTop =2460
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2700
                    ColumnStart =8
                    ColumnEnd =8
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =120
                    Top =120
                    Width =1635
                    Height =2820
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="NotesList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT VendorNotes.VendorNoteID, VendorNotes.vnDate FROM VendorNotes WHERE (((Ve"
                        "ndorNotes.vnVendorID)=[Forms]![VendorsForm]![VendorID]));"
                    ColumnWidths ="0;1440;0"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =1
                    InheritValueList =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =2940
                    RowEnd =4
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "VendorNotes.cls"
