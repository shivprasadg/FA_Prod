Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6864
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =600
    Top =9300
    Right =9960
    Bottom =13425
    RecSrcDt = Begin
        0xc71ea06ed661e540
    End
    RecordSource ="SELECT ScheduleCoLessee.ScheduleID, ScheduleCoLessee.CoLesseeID, ScheduleCoLesse"
        "e.CoLesseeLabel, ScheduleCoLessee.CoLesseeTaxID, ScheduleCoLessee.HideOnCOA, Sch"
        "eduleCoLessee.INDX FROM ScheduleCoLessee;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            AddColon = NotDefault
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
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4230
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2190
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="ScheduleID"
                    ControlSource ="ScheduleID"
                    GridlineColor =0

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2190
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =2400
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1644
                    Top =1200
                    Width =3960
                    Height =312
                    FontSize =10
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =0
                    Name ="CoLesseeTaxID"
                    ControlSource ="CoLesseeTaxID"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =0

                    LayoutCachedLeft =1644
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =1512
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =36
                            Top =1200
                            Width =1536
                            Height =312
                            FontSize =10
                            FontWeight =700
                            BorderColor =0
                            ForeColor =8210719
                            Name ="Label5"
                            Caption ="TaxID"
                            GroupTable =1
                            GridlineColor =0
                            LayoutCachedLeft =36
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1572
                            LayoutCachedHeight =1512
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3696
                    Top =1584
                    Width =1908
                    Height =312
                    FontSize =9
                    TabIndex =5
                    ForeColor =0
                    Name ="SaveCoLessee"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Save this Co Lessee"
                    GroupTable =1
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

                    LayoutCachedLeft =3696
                    LayoutCachedTop =1584
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =1896
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =14136213
                    HoverTint =60.0
                    PressedColor =14136213
                    PressedTint =60.0
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2520
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =8
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="MLID"
                    ControlSource ="INDX"
                    GridlineColor =0

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =2730
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5655
                    Top =2850
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =9
                    BorderColor =12632256
                    ForeColor =5026082
                    Name ="FAID"
                    GridlineColor =0

                    LayoutCachedLeft =5655
                    LayoutCachedTop =2850
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =3060
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5760
                    Left =1644
                    Top =816
                    Width =3960
                    Height =312
                    FontSize =10
                    TabIndex =1
                    BackColor =16249583
                    BorderColor =14277081
                    ForeColor =0
                    Name ="CoLesseeID"
                    ControlSource ="CoLesseeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientID, clientCompanyName AS CoLessee, clientTaxID As TaxID, ClientGrou"
                        "pId FROM  Clients WHERE ClientGroupID= 27 ORDER BY clientCompanyName"
                    ColumnWidths ="0;3168;1152;0"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    GroupTable =1
                    GridlineColor =0

                    LayoutCachedLeft =1644
                    LayoutCachedTop =816
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =1128
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =36
                            Top =816
                            Width =1536
                            Height =312
                            FontSize =10
                            FontWeight =700
                            BorderColor =0
                            ForeColor =8210719
                            Name ="Label29"
                            Caption ="Co/Sub Lessee "
                            GroupTable =1
                            GridlineColor =0
                            LayoutCachedLeft =36
                            LayoutCachedTop =816
                            LayoutCachedWidth =1572
                            LayoutCachedHeight =1128
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1644
                    Top =432
                    Width =3960
                    Height =312
                    FontSize =10
                    BackColor =16249583
                    BorderColor =14277081
                    ForeColor =0
                    Name ="CoLesseeLabel"
                    ControlSource ="CoLesseeLabel"
                    RowSourceType ="Value List"
                    RowSource ="\"Lessee\";\"Co-Lessee\";\"Sublessee\";\"Co-Obligor\""
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    GridlineColor =0

                    LayoutCachedLeft =1644
                    LayoutCachedTop =432
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =744
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =36
                            Top =432
                            Width =1536
                            Height =312
                            FontSize =10
                            FontWeight =700
                            BorderColor =0
                            ForeColor =8210719
                            Name ="Label4"
                            Caption ="Leesee Type "
                            GroupTable =1
                            GridlineColor =0
                            LayoutCachedLeft =36
                            LayoutCachedTop =432
                            LayoutCachedWidth =1572
                            LayoutCachedHeight =744
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =36
                    Top =1584
                    Width =1536
                    Height =312
                    Name ="EmptyCell37"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =1584
                    LayoutCachedWidth =1572
                    LayoutCachedHeight =1896
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =45
                    Top =2010
                    Width =5445
                    Height =1725
                    FontSize =10
                    TabIndex =6
                    Name ="CoLesseeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Sch.INDX, Sch.CoLesseeLabel, Clients.clientCompanyName, Sch.ScheduleId, C"
                        "lients.clientID FROM ScheduleCoLessee AS Sch INNER JOIN Clients ON Sch.CoLesseeI"
                        "D = Clients.clientID WHERE (((Sch.ScheduleId)=1376)) ORDER BY Clients.clientComp"
                        "anyName;"
                    ColumnWidths ="0;1440;3600;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =45
                    LayoutCachedTop =2010
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =3735
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1644
                    Top =1584
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =4
                    ForeColor =0
                    Name ="AddNew"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Add a Co Lessee"
                    GroupTable =1
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000505050ff505050ff00000000505050ff505050ff ,
                        0x00000000505050ff505050ff0000000000000000000000000000000000000000 ,
                        0x000000000000000000000000505050ff505050ff00000000505050ff505050ff ,
                        0x00000000505050ff505050ff0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000072727248 ,
                        0x727272ff727272541c75e3261c75e3ff1c75e326000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727248727272f6 ,
                        0x727272ff727272ff1c75e3ff000000001c75e3ff00000000000000001c75e326 ,
                        0x1c75e3ff1c75e32600000000000000000000000072727248727272f6727272ff ,
                        0x727272fc727272541c75e3261c75e3ff1c75e32600000000000000001c75e3ff ,
                        0x000000001c75e3ff000000000000000072727248727272f6727272ff727272fc ,
                        0x727272540000000000000000000000000000000000000000000000001c75e326 ,
                        0x1c75e3ff1c75e3260000000072727248727272f6727272ff727272fc72727254 ,
                        0x000000000000000000000000000000001c75e3261c75e3ff1c75e32600000000 ,
                        0x000000000000000072727248727272f6727272ff727272fc7272725400000000 ,
                        0x000000000000000000000000000000001c75e3ff000000001c75e3ff00000000 ,
                        0x0000000072727248727272f6727272ff727272f9727272540000000000000000 ,
                        0x000000000000000000000000000000001c75e3261c75e3ff1c75e32600000000 ,
                        0x72727248727272f6727272ff727272f972727254000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000072727248 ,
                        0x727272f6ffffffff727272f97272725400000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727248727272f6 ,
                        0xffffffff727272f9727272540000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x727272f972727254000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727254727272ff ,
                        0x7272725400000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1644
                    LayoutCachedTop =1584
                    LayoutCachedWidth =3624
                    LayoutCachedHeight =1896
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    PressedThemeColorIndex =6
                    PressedTint =40.0
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3495
                    Top =3765
                    Width =1980
                    Height =315
                    FontSize =9
                    TabIndex =10
                    Name ="DeleteMe"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Delete the selected Co-Lessee"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000008080803980808096868686d6 ,
                        0x828282f7808080ff828282f8868686d68585859c808080390000000000000000 ,
                        0x0000000000000000000000000000000000000000808080ff808080ff808080ff ,
                        0x808080ff808080ff808080ff808080ff808080ff808080ff0000000000000000 ,
                        0x0000000000000000000000000000000080808006818181fbcbcbcbffe6e6e6ff ,
                        0xf9f9f9fffffffffff9f9f9ffe6e6e6ffcbcbcbff808080ff0000000000000000 ,
                        0x000000000000000000000000000000008080801e898989edffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff868686fc8080801500000000 ,
                        0x0000000000000000000000000000000080808036939393e6ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff929292fc8080802d00000000 ,
                        0x00000000000000000000000000000000808080519f9f9fe3ffffffffe0e5d4ff ,
                        0x90a468ff90a468ff90a468ffdce2cfffffffffffa0a0a0fe8383834400000000 ,
                        0x0000000000000000000000000000000080808069afafafe5ffffffffadbc8fff ,
                        0xcfd7bdffffffffffd8dfcaffa6b686fffefefdffadadadff9393936a00000000 ,
                        0x0000000000000000000000000000000080808081bebebeecfffffffff6f8f3ff ,
                        0xa0b17dffc2cdacff9eb07cfff3f5efffffffffffb9b9b9ff9999998e00000000 ,
                        0x000000000000000000000000000000008080809ccececef9ffffffffffffffff ,
                        0xeff2eaffbbc7a3ffebefe4ffffffffffffffffffc5c5c5ff979797ad00000000 ,
                        0x00000000000000000000000000000000838383b8ddddddffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffcfcfcfff969696c700000000 ,
                        0x00000000000000000000000000000000848484d3e9e9e9ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffddddddff929292dc00000000 ,
                        0x00000000000000000000000000000000838383edccccccff9b9b9bff808080ff ,
                        0x808080ff808080ff808080ff808080ff9b9b9bffc2c2c2ff8d8d8ded00000000 ,
                        0x00000000000000000000000080808003808080ffa4a4a4ffdededeffffffffff ,
                        0xffffffffffffffffffffffffffffffffdededeffa4a4a4ff808080ff00000000 ,
                        0x00000000000000000000000000000000808080ff9c9c9ccddbdbdbe7ffffffff ,
                        0xffffffffffffffffffffffffffffffffddddddf4a4a4a4dd808080ff00000000 ,
                        0x00000000000000000000000000000000000000008080804e808080bd808080ff ,
                        0x808080ff808080ff808080ff808080ff808080bd8080804e0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3495
                    LayoutCachedTop =3765
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =4080
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =42978
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =1643706
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin Label
                    OverlapFlags =85
                    Width =4455
                    Height =405
                    FontSize =14
                    BorderColor =0
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="Schedule Co/Sub Lessee Manager"
                    FontName ="Segoe UI"
                    GridlineColor =0
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =405
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5676
                    Top =816
                    Width =1152
                    Height =312
                    FontSize =10
                    TabIndex =2
                    BackColor =16249583
                    BorderColor =14277081
                    ForeColor =0
                    Name ="HideOnCOA"
                    ControlSource ="HideOnCOA"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Show\";-1;\"Hide\""
                    ColumnWidths ="0;7200"
                    FontName ="Segoe UI"
                    Format ="True/False"
                    GroupTable =1

                    LayoutCachedLeft =5676
                    LayoutCachedTop =816
                    LayoutCachedWidth =6828
                    LayoutCachedHeight =1128
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5676
                    Top =1200
                    Width =1152
                    Height =312
                    Name ="EmptyCell20"
                    GroupTable =1
                    LayoutCachedLeft =5676
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6828
                    LayoutCachedHeight =1512
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5676
                    Top =1584
                    Width =1152
                    Height =312
                    Name ="EmptyCell21"
                    GroupTable =1
                    LayoutCachedLeft =5676
                    LayoutCachedTop =1584
                    LayoutCachedWidth =6828
                    LayoutCachedHeight =1896
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5676
                    Top =432
                    Width =1152
                    Height =312
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Label0"
                    Caption ="Hide On COA"
                    GroupTable =1
                    LayoutCachedLeft =5676
                    LayoutCachedTop =432
                    LayoutCachedWidth =6828
                    LayoutCachedHeight =744
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
    End
End
CodeBehindForm
' See "Schedule_CoLessee.cls"
