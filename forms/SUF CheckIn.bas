Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17310
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =7755
    Top =2295
    Right =21825
    Bottom =14235
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6482071357eae340
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
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
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
        Begin FormHeader
            Height =525
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =90
                    Width =4110
                    Height =525
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label31"
                    Caption ="SUF Check-in"
                    LayoutCachedLeft =90
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =525
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =14355
                    Top =105
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
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14355
                    LayoutCachedTop =105
                    LayoutCachedWidth =15390
                    LayoutCachedHeight =393
                    PictureCaptionArrangement =5
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13500
                    Top =105
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

                    LayoutCachedLeft =13500
                    LayoutCachedTop =105
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =393
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9900
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =930
                    Width =2265
                    Height =5130
                    Name ="MasterLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLNo FROM MstrLease WHERE (((MstrLease.MLNo) No"
                        "t Like \"demo*\")) ORDER BY MstrLease.MLNo;"
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =930
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =6060
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =30
                            Top =585
                            Width =1335
                            Height =315
                            FontWeight =700
                            Name ="Label27"
                            Caption ="Master Lease"
                            LayoutCachedLeft =30
                            LayoutCachedTop =585
                            LayoutCachedWidth =1365
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =6360
                    Width =2265
                    Height =3450
                    TabIndex =1
                    Name ="SchID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =6360
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =9810
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =30
                            Top =6060
                            Width =1335
                            Height =315
                            FontWeight =700
                            Name ="Label28"
                            Caption ="Schedule"
                            LayoutCachedLeft =30
                            LayoutCachedTop =6060
                            LayoutCachedWidth =1365
                            LayoutCachedHeight =6375
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =30
                    Top =270
                    Width =2265
                    Height =270
                    TabIndex =2
                    Name ="SearchLease"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =270
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =30
                            Top =30
                            Width =2100
                            Height =225
                            FontSize =8
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label29"
                            Caption ="Type SEARCH ▼ HERE"
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =2130
                            LayoutCachedHeight =255
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =2445
                    Top =45
                    Width =14865
                    Height =9750
                    TabIndex =3
                    Name ="Unit_List_SUFSub"
                    SourceObject ="Form.Unit List_SUFSub"
                    LinkChildFields ="SchID"
                    LinkMasterFields ="SchID"
                    VerticalAnchor =2

                    LayoutCachedLeft =2445
                    LayoutCachedTop =45
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =9795
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =2550
                    Top =975
                    Width =2970
                    Height =525
                    FontSize =18
                    FontWeight =700
                    ForeColor =-2147483611
                    Name ="Label30"
                    Caption ="◄ Select a Lease\015\012 "
                    LayoutCachedLeft =2550
                    LayoutCachedTop =975
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1500
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =2550
                    Top =6660
                    Width =3690
                    Height =465
                    FontSize =18
                    FontWeight =700
                    ForeColor =-2147483611
                    Name ="Label32"
                    Caption ="◄Select a Schedule"
                    LayoutCachedLeft =2550
                    LayoutCachedTop =6660
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =7125
                End
            End
        End
        Begin FormFooter
            Height =210
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "SUF CheckIn.cls"
