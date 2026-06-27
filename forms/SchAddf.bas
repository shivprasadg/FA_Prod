Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6675
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =9285
    Top =3735
    Right =19470
    Bottom =8865
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x511ec8152e4ee440
    End
    RecordSource ="SELECT Schedule.SchID, Schedule.SchMLRef, Schedule.SchNo, Schedule.SchShortDesc "
        "FROM Schedule;"
    Caption ="Add a New Schedule"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =348
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Width =4260
                    Height =330
                    FontSize =16
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="Add New Schedule"
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =330
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =5460
                    Top =60
                    Width =1035
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5460
                    LayoutCachedTop =60
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =348
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =4740
                    Top =60
                    Width =720
                    Height =288
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="GoHome"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000adda0fff0bb0fff0da ,
                        0xad0fff0bb0fff0adda0fff0bb0fff0daad0fff0bb0fff0adda0fff0000fff0da ,
                        0xad0ffffffffff0add0b0ffffffff0b0aad0b0ffffff0b0addad0b0ffff0b0ada ,
                        0xadad0b0ff0b00daddadad0b00b010adaadadad0bb0a10daddadadad00ad10ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
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

                    LayoutCachedLeft =4740
                    LayoutCachedTop =60
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =348
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =1755
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5040
                    Left =1530
                    Top =75
                    Width =4020
                    Height =315
                    BackColor =13434879
                    Name ="mlsel"
                    ControlSource ="SchMLRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLOrig, MstrLease.MLNo FROM MstrLease WHERE ((("
                        "MstrLease.Client) Is Not Null)) ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
                    ColumnWidths ="0;0;2160"
                    GroupTable =1

                    LayoutCachedLeft =1530
                    LayoutCachedTop =75
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =45
                            Top =75
                            Width =1425
                            Height =315
                            Name ="Label1"
                            Caption ="Master Lease:"
                            GroupTable =1
                            LayoutCachedLeft =45
                            LayoutCachedTop =75
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =390
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =5490
                    Top =1380
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =1
                    ForeColor =5026082
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =1590
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =450
                    Width =4020
                    Height =315
                    TabIndex =2
                    BackColor =13434879
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    GroupTable =1

                    LayoutCachedLeft =1530
                    LayoutCachedTop =450
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =45
                            Top =450
                            Width =1425
                            Height =315
                            Name ="Label4"
                            Caption ="SchNo:"
                            GroupTable =1
                            LayoutCachedLeft =45
                            LayoutCachedTop =450
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =825
                    Width =4020
                    Height =315
                    ColumnWidth =4335
                    TabIndex =3
                    Name ="SchShortDesc"
                    ControlSource ="SchShortDesc"
                    GroupTable =1

                    LayoutCachedLeft =1530
                    LayoutCachedTop =825
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =1140
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =45
                            Top =825
                            Width =1425
                            Height =315
                            Name ="Label5"
                            Caption ="Desc:"
                            GroupTable =1
                            LayoutCachedLeft =45
                            LayoutCachedTop =825
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =1140
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
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
CodeBehindForm
' See "SchAddf.cls"
