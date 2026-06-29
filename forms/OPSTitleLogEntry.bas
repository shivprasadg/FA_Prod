Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10686
    DatasheetFontHeight =11
    ItemSuffix =228
    Left =5910
    Top =4350
    Right =16800
    Bottom =6900
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf9a49d7613fee440
    End
    RecordSource ="SELECT * FROM UnitTitleLog;"
    BeforeUpdate ="[Event Procedure]"
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
            Height =480
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =75
                    Top =30
                    Width =2265
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Title Log Entry"
                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =840
                    ColumnOrder =1
                    FontSize =9
                    Name ="UnitRef"
                    ControlSource ="UnitRef"

                    LayoutCachedLeft =6540
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8310
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

                    LayoutCachedLeft =8310
                    LayoutCachedTop =60
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =8910
                    Top =60
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    BackStyle =0

                    LayoutCachedLeft =8910
                    LayoutCachedTop =60
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =288
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7440
                    Top =60
                    Width =615
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =3
                    Name ="UnitTitleLogId"
                    ControlSource ="UnitTitleLogId"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =60
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =300
                    BackThemeColorIndex =1
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4635
                    Top =75
                    Height =315
                    TabIndex =4
                    Name ="UserID"
                    ControlSource ="UserID"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =4635
                    LayoutCachedTop =75
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =390
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =2085
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =972
                    Top =36
                    Width =1656
                    Height =312
                    Name ="UnitTitleLogDate"
                    ControlSource ="UnitTitleLogDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Date Required - can be written in"
                    DefaultValue ="Date()"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =972
                    LayoutCachedTop =36
                    LayoutCachedWidth =2628
                    LayoutCachedHeight =348
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =36
                            Width =840
                            Height =312
                            Name ="Label2"
                            Caption ="Date "
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =36
                            LayoutCachedWidth =900
                            LayoutCachedHeight =348
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =972
                    Top =804
                    Width =3300
                    Height =312
                    ColumnWidth =1605
                    TabIndex =2
                    Name ="TitleNumber"
                    ControlSource ="TitleNumber"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =972
                    LayoutCachedTop =804
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1116
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =804
                            Width =840
                            Height =312
                            Name ="Label5"
                            Caption ="Number"
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =804
                            LayoutCachedWidth =900
                            LayoutCachedHeight =1116
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5484
                    Top =420
                    Height =312
                    TabIndex =3
                    Name ="TitleState"
                    ControlSource ="TitleSt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT States.State FROM States ORDER BY States.State;"
                    ColumnWidths ="72"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GroupTable =4

                    LayoutCachedLeft =5484
                    LayoutCachedTop =420
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =732
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4356
                            Top =420
                            Width =1056
                            Height =312
                            Name ="Label10"
                            Caption ="State"
                            GroupTable =4
                            LayoutCachedLeft =4356
                            LayoutCachedTop =420
                            LayoutCachedWidth =5412
                            LayoutCachedHeight =732
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =24
                    ListWidth =5760
                    Left =972
                    Top =420
                    Width =3300
                    Height =312
                    TabIndex =1
                    Name ="EntryReason"
                    ControlSource ="EntryReason"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_TitleReasons.ReasonDescription, LookUp_TitleReasons.ReasonType, Lo"
                        "okUp_TitleReasons.ToReturn, LookUp_TitleReasons.ReasonId FROM LookUp_TitleReason"
                        "s ORDER BY LookUp_TitleReasons.ReasonDescription;"
                    ColumnWidths ="4320;0;0;0"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =972
                    LayoutCachedTop =420
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =732
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =420
                            Width =840
                            Height =312
                            Name ="Label15"
                            Caption ="Reason"
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =900
                            LayoutCachedHeight =732
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4356
                    Top =1572
                    Width =1056
                    Height =300
                    ColumnWidth =1905
                    TabIndex =6
                    Name ="EntryReasonId"
                    ControlSource ="EntryReasonId"
                    GroupTable =4

                    LayoutCachedLeft =4356
                    LayoutCachedTop =1572
                    LayoutCachedWidth =5412
                    LayoutCachedHeight =1872
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5484
                    Top =1572
                    Height =300
                    TabIndex =4
                    Name ="EntryTypeId"
                    ControlSource ="EntryTypeId"
                    GroupTable =4

                    LayoutCachedLeft =5484
                    LayoutCachedTop =1572
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =1872
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5484
                    Top =1188
                    Height =312
                    TabIndex =5
                    Name ="ExpectedReturn"
                    ControlSource ="ExpectedReturn"
                    Format ="True/False"
                    DefaultValue ="-1"
                    GroupTable =4

                    LayoutCachedLeft =5484
                    LayoutCachedTop =1188
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =1500
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4356
                            Top =1188
                            Width =1056
                            Height =312
                            Name ="Label71"
                            Caption ="To Return"
                            GroupTable =4
                            LayoutCachedLeft =4356
                            LayoutCachedTop =1188
                            LayoutCachedWidth =5412
                            LayoutCachedHeight =1500
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6996
                    Top =420
                    Width =2160
                    Height =696
                    FontSize =18
                    FontWeight =700
                    TabIndex =7
                    Name ="SaveMe"
                    Caption ="Save / Close"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6996
                    LayoutCachedTop =420
                    LayoutCachedWidth =9156
                    LayoutCachedHeight =1116
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderThemeColorIndex =0
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    PressedForeThemeColorIndex =0
                    GroupTable =4
                    QuickStyle =22
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =1572
                    Width =4212
                    Height =300
                    FontWeight =700
                    TabIndex =9
                    Name ="AddReason"
                    Caption ="Add A Reason"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =60
                    LayoutCachedTop =1572
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1872
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderThemeColorIndex =0
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    PressedForeThemeColorIndex =0
                    GroupTable =4
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6996
                    Top =1572
                    Width =2160
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    Name ="DeleteLog"
                    Caption ="Delete this entry"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6996
                    LayoutCachedTop =1572
                    LayoutCachedWidth =9156
                    LayoutCachedHeight =1872
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    HoverThemeColorIndex =5
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =4
                    QuickStyle =22
                    QuickStyleMask =-629
                End
                Begin EmptyCell
                    Left =2700
                    Top =36
                    Width =1572
                    Height =312
                    Name ="EmptyCell179"
                    GroupTable =4
                    RightPadding =45
                    LayoutCachedLeft =2700
                    LayoutCachedTop =36
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =348
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =4356
                    Top =36
                    Width =1056
                    Height =312
                    Name ="EmptyCell185"
                    GroupTable =4
                    LayoutCachedLeft =4356
                    LayoutCachedTop =36
                    LayoutCachedWidth =5412
                    LayoutCachedHeight =348
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =60
                    Top =1188
                    Width =4212
                    Height =312
                    Name ="EmptyCell187"
                    GroupTable =4
                    RightPadding =45
                    LayoutCachedLeft =60
                    LayoutCachedTop =1188
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1500
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =6996
                    Top =1188
                    Width =2160
                    Height =312
                    Name ="EmptyCell192"
                    GroupTable =4
                    LayoutCachedLeft =6996
                    LayoutCachedTop =1188
                    LayoutCachedWidth =9156
                    LayoutCachedHeight =1500
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin ComboBox
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5484
                    Top =804
                    Height =312
                    TabIndex =8
                    Name ="InOut"
                    ControlSource ="EntryTypeId"
                    RowSourceType ="Value List"
                    RowSource ="1;\"IN\";2;\"OUT"
                    ColumnWidths ="0;720"
                    GroupTable =4

                    LayoutCachedLeft =5484
                    LayoutCachedTop =804
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =1116
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4356
                            Top =804
                            Width =1056
                            Height =312
                            Name ="Label72"
                            Caption ="Entry Type"
                            GroupTable =4
                            LayoutCachedLeft =4356
                            LayoutCachedTop =804
                            LayoutCachedWidth =5412
                            LayoutCachedHeight =1116
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =4
                        End
                    End
                End
                Begin EmptyCell
                    Left =9216
                    Top =420
                    Height =312
                    Name ="EmptyCell220"
                    GroupTable =4
                    LayoutCachedLeft =9216
                    LayoutCachedTop =420
                    LayoutCachedWidth =10656
                    LayoutCachedHeight =732
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9216
                    Top =804
                    Height =312
                    Name ="EmptyCell221"
                    GroupTable =4
                    LayoutCachedLeft =9216
                    LayoutCachedTop =804
                    LayoutCachedWidth =10656
                    LayoutCachedHeight =1116
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9216
                    Top =1188
                    Height =312
                    Name ="EmptyCell222"
                    GroupTable =4
                    LayoutCachedLeft =9216
                    LayoutCachedTop =1188
                    LayoutCachedWidth =10656
                    LayoutCachedHeight =1500
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9216
                    Top =1572
                    Height =300
                    Name ="EmptyCell223"
                    GroupTable =4
                    LayoutCachedLeft =9216
                    LayoutCachedTop =1572
                    LayoutCachedWidth =10656
                    LayoutCachedHeight =1872
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =5484
                    Top =36
                    Height =312
                    Name ="EmptyCell226"
                    GroupTable =4
                    LayoutCachedLeft =5484
                    LayoutCachedTop =36
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =348
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =6996
                    Top =36
                    Width =3660
                    Height =312
                    Name ="EmptyCell227"
                    GroupTable =4
                    LayoutCachedLeft =6996
                    LayoutCachedTop =36
                    LayoutCachedWidth =10656
                    LayoutCachedHeight =348
                    ColumnStart =5
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
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
' See "OPSTitleLogEntry.cls"
