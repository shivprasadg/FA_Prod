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
    Width =10635
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
                    Left =960
                    Top =30
                    Width =1657
                    Height =315
                    Name ="UnitTitleLogDate"
                    ControlSource ="UnitTitleLogDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Date Required - can be written in"
                    DefaultValue ="Date()"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =960
                    LayoutCachedTop =30
                    LayoutCachedWidth =2617
                    LayoutCachedHeight =345
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
                            Top =30
                            Width =840
                            Height =315
                            Name ="Label2"
                            Caption ="Date "
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =30
                            LayoutCachedWidth =900
                            LayoutCachedHeight =345
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
                    Left =960
                    Top =780
                    Width =3300
                    Height =315
                    ColumnWidth =1605
                    TabIndex =2
                    Name ="TitleNumber"
                    ControlSource ="TitleNumber"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =960
                    LayoutCachedTop =780
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =1095
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
                            Top =780
                            Width =840
                            Height =315
                            Name ="Label5"
                            Caption ="Number"
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =780
                            LayoutCachedWidth =900
                            LayoutCachedHeight =1095
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
                    Left =5445
                    Top =405
                    Height =315
                    TabIndex =3
                    Name ="TitleState"
                    ControlSource ="TitleSt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT States.State FROM States ORDER BY States.State;"
                    ColumnWidths ="72"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GroupTable =4

                    LayoutCachedLeft =5445
                    LayoutCachedTop =405
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =720
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
                            Left =4335
                            Top =405
                            Width =1050
                            Height =315
                            Name ="Label10"
                            Caption ="State"
                            GroupTable =4
                            LayoutCachedLeft =4335
                            LayoutCachedTop =405
                            LayoutCachedWidth =5385
                            LayoutCachedHeight =720
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
                    Left =960
                    Top =405
                    Width =3300
                    Height =315
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

                    LayoutCachedLeft =960
                    LayoutCachedTop =405
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =720
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
                            Top =405
                            Width =840
                            Height =315
                            Name ="Label15"
                            Caption ="Reason"
                            GroupTable =4
                            LayoutCachedLeft =60
                            LayoutCachedTop =405
                            LayoutCachedWidth =900
                            LayoutCachedHeight =720
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
                    Left =4335
                    Top =1530
                    Width =1050
                    Height =300
                    ColumnWidth =1905
                    TabIndex =6
                    Name ="EntryReasonId"
                    ControlSource ="EntryReasonId"
                    GroupTable =4

                    LayoutCachedLeft =4335
                    LayoutCachedTop =1530
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =1830
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
                    Left =5445
                    Top =1530
                    Height =300
                    TabIndex =4
                    Name ="EntryTypeId"
                    ControlSource ="EntryTypeId"
                    GroupTable =4

                    LayoutCachedLeft =5445
                    LayoutCachedTop =1530
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =1830
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
                    Left =5445
                    Top =1155
                    Height =315
                    TabIndex =5
                    Name ="ExpectedReturn"
                    ControlSource ="ExpectedReturn"
                    Format ="True/False"
                    DefaultValue ="-1"
                    GroupTable =4

                    LayoutCachedLeft =5445
                    LayoutCachedTop =1155
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =1470
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
                            Left =4335
                            Top =1155
                            Width =1050
                            Height =315
                            Name ="Label71"
                            Caption ="To Return"
                            GroupTable =4
                            LayoutCachedLeft =4335
                            LayoutCachedTop =1155
                            LayoutCachedWidth =5385
                            LayoutCachedHeight =1470
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
                    Left =6945
                    Top =405
                    Width =2160
                    Height =690
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

                    LayoutCachedLeft =6945
                    LayoutCachedTop =405
                    LayoutCachedWidth =9105
                    LayoutCachedHeight =1095
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
                    Top =1530
                    Width =4200
                    Height =300
                    FontWeight =700
                    TabIndex =9
                    Name ="AddReason"
                    Caption ="Add A Reason"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    RightPadding =45

                    LayoutCachedLeft =60
                    LayoutCachedTop =1530
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =1830
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
                    Left =6945
                    Top =1530
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

                    LayoutCachedLeft =6945
                    LayoutCachedTop =1530
                    LayoutCachedWidth =9105
                    LayoutCachedHeight =1830
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
                    Left =2685
                    Top =30
                    Width =1575
                    Height =315
                    Name ="EmptyCell179"
                    GroupTable =4
                    RightPadding =45
                    LayoutCachedLeft =2685
                    LayoutCachedTop =30
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =345
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =4335
                    Top =30
                    Width =1050
                    Height =315
                    Name ="EmptyCell185"
                    GroupTable =4
                    LayoutCachedLeft =4335
                    LayoutCachedTop =30
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =345
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =60
                    Top =1155
                    Width =4200
                    Height =315
                    Name ="EmptyCell187"
                    GroupTable =4
                    RightPadding =45
                    LayoutCachedLeft =60
                    LayoutCachedTop =1155
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =1470
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =6945
                    Top =1155
                    Width =2160
                    Height =315
                    Name ="EmptyCell192"
                    GroupTable =4
                    LayoutCachedLeft =6945
                    LayoutCachedTop =1155
                    LayoutCachedWidth =9105
                    LayoutCachedHeight =1470
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
                    Left =5445
                    Top =780
                    Height =315
                    TabIndex =8
                    Name ="InOut"
                    ControlSource ="EntryTypeId"
                    RowSourceType ="Value List"
                    RowSource ="1;\"IN\";2;\"OUT"
                    ColumnWidths ="0;720"
                    GroupTable =4

                    LayoutCachedLeft =5445
                    LayoutCachedTop =780
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =1095
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
                            Left =4335
                            Top =780
                            Width =1050
                            Height =315
                            Name ="Label72"
                            Caption ="Entry Type"
                            GroupTable =4
                            LayoutCachedLeft =4335
                            LayoutCachedTop =780
                            LayoutCachedWidth =5385
                            LayoutCachedHeight =1095
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
                    Left =9165
                    Top =405
                    Height =315
                    Name ="EmptyCell220"
                    GroupTable =4
                    LayoutCachedLeft =9165
                    LayoutCachedTop =405
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9165
                    Top =780
                    Height =315
                    Name ="EmptyCell221"
                    GroupTable =4
                    LayoutCachedLeft =9165
                    LayoutCachedTop =780
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9165
                    Top =1155
                    Height =315
                    Name ="EmptyCell222"
                    GroupTable =4
                    LayoutCachedLeft =9165
                    LayoutCachedTop =1155
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =1470
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9165
                    Top =1530
                    Height =300
                    Name ="EmptyCell223"
                    GroupTable =4
                    LayoutCachedLeft =9165
                    LayoutCachedTop =1530
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =1830
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =5445
                    Top =30
                    Height =315
                    Name ="EmptyCell226"
                    GroupTable =4
                    LayoutCachedLeft =5445
                    LayoutCachedTop =30
                    LayoutCachedWidth =6885
                    LayoutCachedHeight =345
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =6945
                    Top =30
                    Width =3660
                    Height =315
                    Name ="EmptyCell227"
                    GroupTable =4
                    LayoutCachedLeft =6945
                    LayoutCachedTop =30
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =345
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
