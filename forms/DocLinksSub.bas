Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =449
    Left =12360
    Top =5190
    Right =25320
    Bottom =15780
    ShortcutMenuBar ="SortClip"
    OrderBy ="[DocLinksSub].[dType], [DocLinksSub].[SGrpNO]"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xdac745a792b4e540
    End
    RecordSource ="SELECT \"\" & Chr(13) & Chr(10) & IIf([dtype]=\"DIR\",\"All\",[G].[SGrpNO]) AS s"
        "GrpNo, DocLinks.* FROM DocLinks LEFT JOIN (Select sGrpNo,SGrpId from SchGrp)  AS"
        " G ON DocLinks.GroupID = G.SGrpID;"
    Caption ="DocLinksSub"
    OnCurrent ="[Event Procedure]"
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
            Height =435
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =225
                    Width =12540
                    Height =210
                    BorderColor =15527148
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =225
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =435
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =4200
                    Top =30
                    Width =2460
                    Height =270
                    FontSize =9
                    Name ="Dir_Tags"
                    Caption ="Dir - Tags"
                    GroupTable =6
                    LayoutCachedLeft =4200
                    LayoutCachedTop =30
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =300
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =840
                    Top =30
                    Width =2475
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitPlateNum_Label"
                    Caption ="Doc or Directory Name"
                    GroupTable =6
                    LayoutCachedLeft =840
                    LayoutCachedTop =30
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =6720
                    Top =30
                    Width =1920
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitYr_Label"
                    Caption =" Description"
                    GroupTable =6
                    LayoutCachedLeft =6720
                    LayoutCachedTop =30
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =300
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    Left =8700
                    Top =30
                    Width =3855
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label86"
                    Caption ="Link"
                    GroupTable =6
                    LayoutCachedLeft =8700
                    LayoutCachedTop =30
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =300
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =3375
                    Top =30
                    Width =765
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="unitinserv_Label"
                    Caption ="Type"
                    GroupTable =6
                    LayoutCachedLeft =3375
                    LayoutCachedTop =30
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =300
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12615
                    Top =30
                    Width =285
                    Height =270
                    Name ="EmptyCell417"
                    GroupTable =6
                    LayoutCachedLeft =12615
                    LayoutCachedTop =30
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =300
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =90
                    Top =30
                    Width =690
                    Height =270
                    FontSize =9
                    Name ="Label445"
                    Caption ="Group"
                    GroupTable =6
                    LayoutCachedLeft =90
                    LayoutCachedTop =30
                    LayoutCachedWidth =780
                    LayoutCachedHeight =300
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =779
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =9
                    Name ="DocID"
                    ControlSource ="DocID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8700
                    Top =30
                    Width =3855
                    Height =675
                    FontSize =9
                    TabIndex =5
                    BorderColor =12765388
                    Name ="dLink"
                    ControlSource ="dLink"
                    GroupTable =6

                    LayoutCachedLeft =8700
                    LayoutCachedTop =30
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =30
                    Width =1920
                    Height =675
                    ColumnWidth =990
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =4
                    Name ="dDesc"
                    ControlSource ="dDesc"
                    GroupTable =6

                    LayoutCachedLeft =6720
                    LayoutCachedTop =30
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =30
                    Width =2460
                    Height =675
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =3
                    Name ="dTags"
                    ControlSource ="dTags"
                    GroupTable =6

                    LayoutCachedLeft =4200
                    LayoutCachedTop =30
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Top =720
                    Width =12540
                    Height =59
                    FontSize =10
                    TabIndex =14
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000100000000000000000000001500000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0044006f006300490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d0000000000
                    End
                    HorizontalAnchor =2
                    VerticalAnchor =1

                    LayoutCachedTop =720
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =779
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00140000005b00 ,
                        0x44006f006300490044005d003d005b0074007800740043007500720072006500 ,
                        0x6e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3375
                    Top =30
                    Width =765
                    Height =675
                    ColumnWidth =1275
                    FontWeight =700
                    TabIndex =2
                    BorderColor =12765388
                    Name ="dType"
                    ControlSource ="dType"
                    ConditionalFormat = Begin
                        0x010000008c000000030000000000000002000000000000000700000001010000 ,
                        0xffffff00758c48000000000002000000080000000e00000001010000ffffff00 ,
                        0xcc66000000000000020000000f0000001500000001010000ffffff00cc990000 ,
                        0x220078006c007300780022000000000022007000640066002200000000002200 ,
                        0x440049005200220000000000
                    End
                    GroupTable =6

                    LayoutCachedLeft =3375
                    LayoutCachedTop =30
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    ForeThemeColorIndex =1
                    ConditionalFormat14 = Begin
                        0x010003000000000000000200000001010000ffffff00758c4800060000002200 ,
                        0x78006c0073007800220000000000000000000000000000000000000000000000 ,
                        0x0000000200000001010000ffffff00cc66000005000000220070006400660022 ,
                        0x0000000000000000000000000000000000000000000000000000020000000101 ,
                        0x0000ffffff00cc99000005000000220044004900520022000000000000000000 ,
                        0x00000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =12615
                    Top =435
                    Width =285
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="DeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =6
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000696969126969698a69696951000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000069696906 ,
                        0x696969a5696969276969696f696969ff696969fc6969694e0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000069696906696969b7 ,
                        0x696969b40000000069696921696969f9696969ff696969f66969693000000000 ,
                        0x000000000000000000000000000000000000000069696906696969b4696969f9 ,
                        0x6969692400000000000000006969697e696969ff696969ff696969e46969691b ,
                        0x0000000000000000000000000000000069696906696969ae696969ff69696984 ,
                        0x00000000000000000000000000000000696969a5696969ff696969ff696969d2 ,
                        0x6969690f000000000000000069696906696969ae696969ff696969d269696906 ,
                        0x0000000000000000000000000000000069696906696969b4696969ff696969ff ,
                        0x696969c66969691269696906696969ab696969ff696969f96969693000000000 ,
                        0x00000000000000000000000000000000000000006969690369696993696969ff ,
                        0x696969ff696969d8696969c0696969ff696969ff696969630000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969695d ,
                        0x696969f3696969ff696969ff696969ff6969699c000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969694e ,
                        0x696969f3696969ff696969ff696969ff69696951000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006969690c69696999696969ff ,
                        0x696969ff696969ff696969cc696969e1696969ff6969698a6969690600000000 ,
                        0x0000000000000000000000000000000069696948696969e1696969ff696969ff ,
                        0x696969ff696969ab696969096969690969696984696969f6696969cc69696921 ,
                        0x0000000000000000000000006969698d696969ff696969ff696969ff696969ff ,
                        0x69696987000000000000000000000000000000006969691e69696993696969ea ,
                        0x696969786969690669696903696969f6696969ff696969ff696969d569696942 ,
                        0x0000000000000000000000000000000000000000000000000000000069696912 ,
                        0x6969696369696927000000006969694e696969b76969696f6969690600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12615
                    LayoutCachedTop =435
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =705
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =6
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =30
                    Width =2475
                    Height =675
                    FontSize =9
                    TabIndex =1
                    Name ="dName"
                    ControlSource ="dName"
                    GroupTable =6

                    LayoutCachedLeft =840
                    LayoutCachedTop =30
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =10
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =8
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =11
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =12
                    Name ="GroupID"
                    ControlSource ="GroupID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10035
                    Top =30
                    Width =0
                    Height =570
                    FontSize =9
                    TabIndex =13
                    Name ="UnitID"
                    ControlSource ="UnitID"

                    LayoutCachedLeft =10035
                    LayoutCachedTop =30
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =12615
                    Top =30
                    Width =285
                    Height =345
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="EditLink"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =6
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12615
                    LayoutCachedTop =30
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =375
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =6
                    QuickStyle =22
                    QuickStyleMask =-117
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =90
                    Top =30
                    Width =690
                    Height =675
                    FontWeight =700
                    Name ="Text440"
                    ControlSource ="sGrpNo"
                    GroupTable =6

                    LayoutCachedLeft =90
                    LayoutCachedTop =30
                    LayoutCachedWidth =780
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
            End
        End
        Begin FormFooter
            Height =240
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =240
                    Width =1979
                    FontSize =8
                    BorderColor =12632256
                    Name ="UnitCount"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =2219
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3180
                    Width =765
                    Height =210
                    FontSize =8
                    TabIndex =1
                    BorderColor =12632256
                    Name ="txtCurrent"

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =3945
                    LayoutCachedHeight =210
                End
            End
        End
    End
End
CodeBehindForm
' See "DocLinksSub.cls"
