Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8064
    DatasheetFontHeight =11
    ItemSuffix =21
    Left =9885
    Top =2415
    Right =18210
    Bottom =6795
    Filter ="INDx =1304"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x007d979177eae440
    End
    RecordSource ="ATF_ChangeReasons"
    Caption ="Edit Reason Note"
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            Height =1455
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureAlignment =0
                    PictureType =2
                    Top =1260
                    Width =8064
                    Height =195
                    Name ="Image439"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =1260
                    LayoutCachedWidth =8064
                    LayoutCachedHeight =1455
                    TabIndex =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Top =360
                    Width =660
                    Height =300
                    ForeColor =8872517
                    Name ="Label2"
                    Caption ="Action"
                    RightPadding =38
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =720
                    LayoutCachedHeight =660
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =750
                    Top =360
                    Width =7125
                    Height =315
                    TabIndex =1
                    Name ="ReasonType"
                    ControlSource ="Action"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =750
                    LayoutCachedTop =360
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =675
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =7845
                    Height =360
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =8872517
                    Name ="Header"
                    ControlSource ="=\"Edit ATF Change Reason For Sch: \" & [SchNo] & \" Grp: \" & [GroupNo]"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =750
                    Top =600
                    Width =7125
                    FontSize =10
                    Name ="info"
                    ControlSource ="=\"Date Changed: <b>\" & [ChangeDate]"
                    RightPadding =38
                    BottomPadding =38
                    TextFormat =1

                    LayoutCachedLeft =750
                    LayoutCachedTop =600
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =840
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    BackStyle =0
                    IMESentenceMode =3
                    Left =750
                    Top =840
                    Width =7125
                    FontSize =10
                    TabIndex =3
                    Name ="Text19"
                    ControlSource ="=\"Prev:<b> \" & [OldValue]"
                    RightPadding =38
                    BottomPadding =38
                    TextFormat =1

                    LayoutCachedLeft =750
                    LayoutCachedTop =840
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =1080
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =750
                    Top =1080
                    Width =7125
                    FontSize =10
                    TabIndex =4
                    Name ="Text20"
                    ControlSource ="=\"Curr:<b> \" & [NewValue]"
                    RightPadding =38
                    BottomPadding =38
                    TextFormat =1

                    LayoutCachedLeft =750
                    LayoutCachedTop =1080
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =1320
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
            End
        End
        Begin Section
            Height =2940
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =135
                    Top =285
                    Width =7845
                    Height =2175
                    ColumnWidth =3465
                    Name ="ReasonDescription"
                    ControlSource ="Reason"
                    RightPadding =38
                    BottomPadding =38
                    TextFormat =1

                    LayoutCachedLeft =135
                    LayoutCachedTop =285
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =2460
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    Left =135
                    Width =3975
                    Height =300
                    ForeColor =8872517
                    Name ="Label1"
                    Caption ="Reason/Note"
                    RightPadding =38
                    LayoutCachedLeft =135
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =135
                    Top =2520
                    Width =1680
                    Height =278
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8872517
                    Name ="Undo"
                    Caption ="Undo"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =135
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =2798
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
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =6300
                    Top =2520
                    Width =1680
                    Height =278
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8872517
                    Name ="SaveChanges"
                    Caption ="Save Changes"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =6300
                    LayoutCachedTop =2520
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =2798
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
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2685
                    Top =2550
                    Width =2400
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4754549
                    ForeColor =10855845
                    Name ="INDX"
                    ControlSource ="INDX"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2685
                    LayoutCachedTop =2550
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =2790
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
' See "ATFReasonEdit.cls"
