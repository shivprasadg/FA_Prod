Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14688
    DatasheetFontHeight =11
    ItemSuffix =440
    Left =5670
    Top =5640
    Right =20610
    Bottom =11160
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0ababd44b1a0e540
    End
    RecordSource ="Select * From CloneSchGrp_WorkTable"
    Caption ="Clone Schedule and Groups"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    WaitForPostProcessing =255
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1920
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    SizeMode =0
                    PictureAlignment =0
                    PictureType =2
                    Top =300
                    Width =14688
                    Height =195
                    Name ="Image439"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =300
                    LayoutCachedWidth =14688
                    LayoutCachedHeight =495
                    TabIndex =7
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =7845
                    Height =360
                    FontSize =14
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8872517
                    Name ="Header"
                    ControlSource ="=\"Schedule | Group Cloning\""

                    LayoutCachedLeft =60
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =216
                    Top =1308
                    Width =1476
                    Height =588
                    BorderColor =0
                    Name ="Label26"
                    Caption ="Source\015\012Schedule Num"
                    GroupTable =8
                    LayoutCachedLeft =216
                    LayoutCachedTop =1308
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =1896
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1764
                    Top =1308
                    Width =1608
                    Height =588
                    BorderColor =0
                    Name ="Label27"
                    Caption ="Source\015\012Group Num"
                    GroupTable =8
                    LayoutCachedLeft =1764
                    LayoutCachedTop =1308
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =1896
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4188
                    Top =1308
                    Width =2520
                    Height =588
                    BorderColor =0
                    Name ="Label32"
                    Caption ="New \015\012Group Num"
                    GroupTable =8
                    LayoutCachedLeft =4188
                    LayoutCachedTop =1308
                    LayoutCachedWidth =6708
                    LayoutCachedHeight =1896
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =6780
                    Top =1308
                    Width =4740
                    Height =588
                    BorderColor =0
                    Name ="Label33"
                    Caption ="New \015\012Group Description"
                    GroupTable =8
                    LayoutCachedLeft =6780
                    LayoutCachedTop =1308
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1896
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3444
                    Top =1308
                    Width =672
                    Height =588
                    BorderColor =0
                    Name ="Label28"
                    Caption ="Clone \015\012Me"
                    GroupTable =8
                    LayoutCachedLeft =3444
                    LayoutCachedTop =1308
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =1896
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =6780
                    Top =912
                    Width =4740
                    Height =336
                    TabIndex =4
                    BorderColor =8210719
                    Name ="NewSchDesctoUse"
                    ControlSource ="NewSchDesc"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =8

                    LayoutCachedLeft =6780
                    LayoutCachedTop =912
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =216
                    Top =912
                    Width =1476
                    Height =336
                    Name ="EmptyCell393"
                    GroupTable =8
                    LayoutCachedLeft =216
                    LayoutCachedTop =912
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =4188
                    Top =912
                    Width =2520
                    Height =336
                    TabIndex =3
                    BorderColor =8210719
                    Name ="NewSchNotoUse"
                    ControlSource ="NewSchNo"
                    AfterUpdate ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    GroupTable =8

                    LayoutCachedLeft =4188
                    LayoutCachedTop =912
                    LayoutCachedWidth =6708
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =119
                    TextAlign =1
                    Left =6780
                    Top =480
                    Width =4740
                    Height =360
                    BorderColor =0
                    Name ="Label404"
                    Caption ="▼New Sch Description to use ▼"
                    GroupTable =8
                    LayoutCachedLeft =6780
                    LayoutCachedTop =480
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =840
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =216
                    Top =480
                    Width =1476
                    Height =360
                    Name ="EmptyCell408"
                    GroupTable =8
                    LayoutCachedLeft =216
                    LayoutCachedTop =480
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =840
                    LayoutGroup =1
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =127
                    Left =4188
                    Top =480
                    Width =2520
                    Height =360
                    BorderColor =0
                    Name ="Label416"
                    Caption ="▼New Schdule Num ▼"
                    GroupTable =8
                    LayoutCachedLeft =4188
                    LayoutCachedTop =480
                    LayoutCachedWidth =6708
                    LayoutCachedHeight =840
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1764
                    Top =912
                    Width =1608
                    Height =336
                    TabIndex =2
                    BorderColor =12632256
                    ForeColor =9211020
                    Name ="NewSchIDtoUse"
                    GroupTable =8

                    LayoutCachedLeft =1764
                    LayoutCachedTop =912
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =127
                    TextFontCharSet =0
                    Left =1764
                    Top =480
                    Width =2352
                    FontWeight =700
                    TabIndex =1
                    ForeColor =3751056
                    Name ="CloneSchNGroup"
                    Caption ="Clone  Selected   ▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Insert this NEW Group"
                    GroupTable =8

                    LayoutCachedLeft =1764
                    LayoutCachedTop =480
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =4
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedTint =60.0
                    HoverForeColor =3751056
                    PressedForeColor =3751056
                    GroupTable =8
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =3444
                    Top =912
                    Width =672
                    Height =336
                    Name ="EmptyCell437"
                    GroupTable =8
                    LayoutCachedLeft =3444
                    LayoutCachedTop =912
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =1248
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =8
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =11880
                    Top =1200
                    Width =2175
                    Height =585
                    BorderColor =0
                    Name ="Label439"
                    Caption =" Copy a NEW Group to\015\012 the SAME Schedule ▼"
                    GroupTable =10
                    LayoutCachedLeft =11880
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14055
                    LayoutCachedHeight =1785
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                End
                Begin CommandButton
                    OverlapFlags =127
                    TextFontCharSet =0
                    Left =13140
                    Top =45
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="GoHome"
                    Caption ="Home"
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

                    LayoutCachedLeft =13140
                    LayoutCachedTop =45
                    LayoutCachedWidth =13695
                    LayoutCachedHeight =273
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
                    OverlapFlags =247
                    TextFontCharSet =0
                    PictureType =1
                    Left =13707
                    Top =45
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =13707
                    LayoutCachedTop =45
                    LayoutCachedWidth =14262
                    LayoutCachedHeight =273
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
            Height =420
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Width =4140
                    Height =420
                    Name ="Box436"
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =420
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =216
                    Top =36
                    Width =1476
                    Height =336
                    BorderColor =12632256
                    Name ="SourceSchNo"
                    ControlSource ="SourceSchNo"
                    GroupTable =8

                    LayoutCachedLeft =216
                    LayoutCachedTop =36
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =372
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1764
                    Top =36
                    Width =1608
                    Height =336
                    TabIndex =1
                    BorderColor =12632256
                    Name ="SourceGroupNo"
                    ControlSource ="SourceGroupNo"
                    GroupTable =8

                    LayoutCachedLeft =1764
                    LayoutCachedTop =36
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =372
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =4188
                    Top =36
                    Width =2520
                    Height =336
                    TabIndex =3
                    BorderColor =8210719
                    Name ="NewGroupNo"
                    ControlSource ="NewGroupNo"
                    GroupTable =8

                    LayoutCachedLeft =4188
                    LayoutCachedTop =36
                    LayoutCachedWidth =6708
                    LayoutCachedHeight =372
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =6780
                    Top =36
                    Width =4740
                    Height =336
                    TabIndex =4
                    BorderColor =8210719
                    Name ="NewGroupDesc"
                    ControlSource ="NewGroupDesc"
                    GroupTable =8

                    LayoutCachedLeft =6780
                    LayoutCachedTop =36
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =372
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin ToggleButton
                    OverlapFlags =215
                    Left =3444
                    Top =36
                    Width =672
                    Height =336
                    TabIndex =2
                    Name ="CloneMe"
                    ControlSource ="CloneMe"
                    DefaultValue ="0"
                    Caption ="No"
                    GroupTable =8

                    LayoutCachedLeft =3444
                    LayoutCachedTop =36
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =372
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedColor =4754549
                    PressedForeColor =4754549
                    GroupTable =8
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =11580
                    Top =36
                    Width =3000
                    Height =336
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="JustCopyGroup"
                    Caption ="◄ Copy THIS Group to SAME Sch"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Insert this NEW Group"
                    GroupTable =9

                    LayoutCachedLeft =11580
                    LayoutCachedTop =36
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =372
                    PictureCaptionArrangement =4
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =16777215
                    HoverColor =4754549
                    PressedColor =14211288
                    PressedForeThemeColorIndex =6
                    PressedForeShade =50.0
                    GroupTable =9
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                End
            End
        End
        Begin FormFooter
            Height =278
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =9360
                    Width =1680
                    Height =278
                    FontSize =8
                    FontWeight =700
                    ForeColor =8872517
                    Name ="SaveChanges"
                    Caption ="Save Changes"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff00000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xd5abc3ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff00000000727272ff727272ff727272ff727272ff00000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff
                    End

                    LayoutCachedLeft =9360
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =278
                    PictureCaptionArrangement =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverColor =12566463
                    HoverForeColor =8872517
                    PressedForeColor =8872517
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Width =2400
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="INDX"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =240
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =180
                    Width =1680
                    Height =278
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8872517
                    Name ="Undo"
                    Caption ="Undo"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4aedb17d4ac0b17d4a7bb17d4a0c00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4ae7b17d4a48 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a0fb17d4a3fb17d4aa8b17d4affb17d4af9 ,
                        0xb17d4a3000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a96b17d4aff ,
                        0xb17d4ab100000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a12b17d4aff ,
                        0xb17d4af000000000000000000000000000000000b17d4a5ab17d4afcb17d4aff ,
                        0xb17d4af9b17d4a4500000000000000000000000000000000b17d4a12b17d4aff ,
                        0xb17d4af6000000000000000000000000b17d4a42b17d4af9b17d4affb17d4afc ,
                        0xb17d4a510000000000000000000000000000000000000000b17d4a96b17d4aff ,
                        0xb17d4ac30000000000000000b17d4a36b17d4af6b17d4affb17d4affb17d4a5d ,
                        0x000000000000000000000000b17d4a12b17d4a42b17d4aa8b17d4affb17d4aff ,
                        0xb17d4a4b00000000b17d4a27b17d4aeab17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4af9b17d4a78 ,
                        0x0000000000000000b17d4a24b17d4aeab17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aedb17d4accb17d4a90b17d4a2400000000 ,
                        0x000000000000000000000000b17d4a2db17d4aedb17d4affb17d4affb17d4a5a ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000b17d4a36b17d4af3b17d4affb17d4af9 ,
                        0xb17d4a3c00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a42b17d4af6b17d4aff ,
                        0xb17d4aeab17d4a24000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =278
                    PictureCaptionArrangement =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverColor =12566463
                    HoverForeColor =8872517
                    PressedForeColor =8872517
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Width =660
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="SchID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5280
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =240
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8340
                    Width =720
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="Text22"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8340
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =240
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
    End
End
CodeBehindForm
' See "CloneScheduleGroups.cls"
