Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10845
    DatasheetFontHeight =11
    ItemSuffix =47
    Left =-13560
    Top =9255
    Right =-2460
    Bottom =11280
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0f57a628f961e440
    End
    RecordSource ="SELECT UnitTitleLog.* FROM UnitTitleLog;"
    Caption ="Input A Title Log"
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
            Height =495
            BackColor =15527148
            Name ="FormHeader"
            AutoHeight =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =30
                    Width =5265
                    Height =465
                    FontSize =16
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Make Title Log Entry"
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =495
                End
            End
        End
        Begin Section
            Height =1545
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9885
                    Top =1155
                    Width =780
                    FontSize =9
                    TabIndex =5
                    Name ="UnitRef"
                    ControlSource ="UnitRef"

                    LayoutCachedLeft =9885
                    LayoutCachedTop =1155
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1404
                    Top =108
                    Width =1332
                    Height =300
                    Name ="UnitTitleLogDate"
                    ControlSource ="UnitTitleLogDate"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Date Required - can be written in"
                    DefaultValue ="=Date()"
                    InputMask ="99/99/00;0;"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1404
                    LayoutCachedTop =108
                    LayoutCachedWidth =2736
                    LayoutCachedHeight =408
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =252
                            Top =108
                            Width =1080
                            Height =300
                            Name ="Label2"
                            Caption ="Date:"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =252
                            LayoutCachedTop =108
                            LayoutCachedWidth =1332
                            LayoutCachedHeight =408
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6372
                    Top =600
                    Width =3432
                    Height =312
                    ColumnWidth =1605
                    TabIndex =4
                    Name ="TitleNumber"
                    ControlSource ="TitleNumber"
                    GroupTable =1
                    BottomPadding =150
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="MsgBox"
                            Argument ="Title Number does not have to be entered when title is going out. Enter when tit"
                                "le coming in."
                            Argument ="-1"
                            Argument ="4"
                            Argument ="When is Title No. Required?"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"TitleNumber\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"MessageBox\"><Argu"
                                "ment Name=\"Message\">Title Numbe"
                        End
                        Begin
                            Comment ="_AXL:r does not have to be entered when title is going out. Enter when title com"
                                "ing in.</Argument><Argument Name=\"Type\">Information</Argument><Argument Name=\""
                                "Title\">When is Title No. Required?</Argument></Action></Statements></UserInterf"
                                "aceMacro>"
                        End
                    End

                    LayoutCachedLeft =6372
                    LayoutCachedTop =600
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =912
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4932
                            Top =600
                            Width =1368
                            Height =312
                            Name ="Label5"
                            Caption ="Title Number:"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =4932
                            LayoutCachedTop =600
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =912
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    Left =4020
                    Top =108
                    Width =852
                    Height =300
                    TabIndex =1
                    Name ="EntryTypeId"
                    ControlSource ="EntryTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EntryType.EntryTypeId, EntryType.EntryTypeDescription FROM EntryType ORDE"
                        "R BY EntryType.EntryTypeId;"
                    ColumnWidths ="0;720"
                    GroupTable =1
                    BottomPadding =150
                    AfterUpdateEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="EntryReasonId"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"EntryTypeId\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"Requery\"><Argumen"
                                "t Name=\"ControlName\">EntryReaso"
                        End
                        Begin
                            Comment ="_AXL:nId</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4020
                    LayoutCachedTop =108
                    LayoutCachedWidth =4872
                    LayoutCachedHeight =408
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2796
                            Top =108
                            Width =1152
                            Height =300
                            Name ="Label3"
                            Caption ="Entry Type :"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =2796
                            LayoutCachedTop =108
                            LayoutCachedWidth =3948
                            LayoutCachedHeight =408
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =1404
                    Top =600
                    Width =1332
                    Height =312
                    Name ="EmptyCell42"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =1404
                    LayoutCachedTop =600
                    LayoutCachedWidth =2736
                    LayoutCachedHeight =912
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =252
                    Top =600
                    Width =1080
                    Height =312
                    Name ="EmptyCell43"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =252
                    LayoutCachedTop =600
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =912
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4020
                    Top =600
                    Width =852
                    Height =312
                    TabIndex =3
                    Name ="TitleSt"
                    ControlSource ="TitleSt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_States.StateAbbrev AS State FROM LookUp_States ORDER BY LookUp_Sta"
                        "tes.StateAbbrev;"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =4020
                    LayoutCachedTop =600
                    LayoutCachedWidth =4872
                    LayoutCachedHeight =912
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2796
                            Top =600
                            Width =1152
                            Height =312
                            Name ="Label6"
                            Caption ="Title State:"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =2796
                            LayoutCachedTop =600
                            LayoutCachedWidth =3948
                            LayoutCachedHeight =912
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6372
                    Top =108
                    Width =3432
                    Height =300
                    TabIndex =2
                    Name ="EntryReasonId"
                    ControlSource ="EntryReasonId"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =6372
                    LayoutCachedTop =108
                    LayoutCachedWidth =9804
                    LayoutCachedHeight =408
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4932
                            Top =108
                            Width =1368
                            Height =300
                            Name ="Label4"
                            Caption ="Reason:"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =4932
                            LayoutCachedTop =108
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =408
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
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
