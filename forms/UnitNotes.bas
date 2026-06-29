Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9255
    DatasheetFontHeight =11
    ItemSuffix =28
    Left =2895
    Top =11700
    Right =12300
    Bottom =18060
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x79780acbd297e540
    End
    OnDirty ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            Height =4095
            Name ="Detail"
            AlternateBackColor =14545355
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2340
                    Top =108
                    Width =1368
                    Height =276
                    TabIndex =2
                    BackColor =11271167
                    Name ="uNoteDate"
                    ControlSource ="uNoteDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Note Date is required"
                    DefaultValue ="=Date()"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =2340
                    LayoutCachedTop =108
                    LayoutCachedWidth =3708
                    LayoutCachedHeight =384
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1836
                            Top =108
                            Width =432
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="Label0"
                            Caption ="Date"
                            GroupTable =2
                            BottomPadding =150
                            LayoutCachedLeft =1836
                            LayoutCachedTop =108
                            LayoutCachedWidth =2268
                            LayoutCachedHeight =384
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2160
                    Top =480
                    Width =6915
                    Height =3498
                    TabIndex =6
                    Name ="uNote"
                    ControlSource ="uNote"
                    OnLostFocus ="[Event Procedure]"
                    TextFormat =1

                    LayoutCachedLeft =2160
                    LayoutCachedTop =480
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =3978
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8520
                    Top =108
                    Width =528
                    Height =276
                    TabIndex =5
                    Name ="UnitRef"
                    ControlSource ="UnitRef"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =8520
                    LayoutCachedTop =108
                    LayoutCachedWidth =9048
                    LayoutCachedHeight =384
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =5652
                    Top =108
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="AddUnitNote"
                    Caption ="Add Note"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =150

                    LayoutCachedLeft =5652
                    LayoutCachedTop =108
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =384
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
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
                    GroupTable =2
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
                    Left =1836
                    Top =492
                    Width =288
                    Height =1143
                    FontSize =8
                    Name ="Label1744"
                    Caption =" └─ Notes"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1836
                    LayoutCachedTop =492
                    LayoutCachedWidth =2124
                    LayoutCachedHeight =1635
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =75
                    Top =90
                    Width =1635
                    Height =3885
                    FontSize =10
                    FontWeight =700
                    Name ="NotesList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =90
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =3975
                    RowEnd =4
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7200
                    Top =108
                    Width =1248
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =7200
                    LayoutCachedTop =108
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =384
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
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
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4116
                    Top =108
                    Height =276
                    TabIndex =1
                    BackColor =11271167
                    Name ="uEmp"
                    ControlSource ="uEmp"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Intials of person entering note are required"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =4116
                    LayoutCachedTop =108
                    LayoutCachedWidth =5556
                    LayoutCachedHeight =384
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3768
                            Top =108
                            Width =276
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="Label10"
                            Caption ="By"
                            GroupTable =2
                            BottomPadding =150
                            LayoutCachedLeft =3768
                            LayoutCachedTop =108
                            LayoutCachedWidth =4044
                            LayoutCachedHeight =384
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
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
' See "UnitNotes.cls"
