Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =11
    ItemSuffix =77
    Left =4110
    Top =3315
    Right =14535
    Bottom =9585
    RecSrcDt = Begin
        0xfcdb78321ba1e540
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CustomControl
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =570
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Width =3825
                    Height =525
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Add / Edit Linked Doc"
                    FontName ="Segoe UI"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3945
                    LayoutCachedHeight =525
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =420
                    Width =8640
                    Height =150
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedTop =420
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =570
                    TabIndex =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7155
                    Top =60
                    Width =720
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="DocID"
                    ControlSource ="DocID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7155
                    LayoutCachedTop =60
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6660
                            Top =60
                            Width =495
                            Height =240
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label74"
                            Caption ="DocID"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =6660
                            LayoutCachedTop =60
                            LayoutCachedWidth =7155
                            LayoutCachedHeight =300
                        End
                    End
                End
            End
        End
        Begin Section
            Height =3135
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =7500
                    Height =585
                    ColumnOrder =0
                    TabIndex =1
                    ForeColor =1279872587
                    Name ="dLink"
                    ControlSource ="dLink"
                    StatusBarText ="hyperlink to Document shown here"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =645
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Top =60
                            Width =1020
                            Height =585
                            Name ="Label11"
                            Caption ="Paste Link HERE "
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedTop =60
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =645
                            LayoutGroup =1
                            BackThemeColorIndex =6
                            BackShade =75.0
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =645
                    Width =7500
                    Height =315
                    ColumnOrder =2
                    BorderColor =10921638
                    Name ="dName"
                    ControlSource ="dName"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =645
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =960
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Top =645
                            Width =1020
                            Height =315
                            Name ="Label0"
                            Caption ="FileName"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedTop =645
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =960
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =960
                    Width =7500
                    Height =315
                    ColumnOrder =3
                    TabIndex =2
                    BorderColor =10921638
                    Name ="dtype"
                    ControlSource ="dtype"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =960
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1275
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Top =960
                            Width =1020
                            Height =315
                            Name ="Label1"
                            Caption ="Doc Type"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedTop =960
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =1275
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =1275
                    Width =7500
                    Height =675
                    ColumnOrder =4
                    TabIndex =3
                    BorderColor =10921638
                    Name ="dtags"
                    ControlSource ="dtags"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1275
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1950
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Top =1275
                            Width =1020
                            Height =675
                            Name ="Label2"
                            Caption ="Dir / Tags"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedTop =1275
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =1950
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =855
                    Top =2580
                    Width =619
                    Height =195
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ClientID"
                    ControlSource ="ClientID"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =855
                    LayoutCachedTop =2580
                    LayoutCachedWidth =1474
                    LayoutCachedHeight =2775
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =240
                            Top =2580
                            Width =619
                            Height =195
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label9"
                            Caption ="ClientID"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =2580
                            LayoutCachedWidth =859
                            LayoutCachedHeight =2775
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3315
                    Top =2580
                    Width =615
                    Height =195
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SchID"
                    ControlSource ="SchID"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3315
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =2775
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2700
                            Top =2580
                            Width =615
                            Height =195
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label10"
                            Caption ="SchID"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2700
                            LayoutCachedTop =2580
                            LayoutCachedWidth =3315
                            LayoutCachedHeight =2775
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2085
                    Top =2580
                    Width =615
                    Height =195
                    ColumnOrder =7
                    FontSize =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="MLID"
                    ControlSource ="MLID"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2085
                    LayoutCachedTop =2580
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =2775
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =1470
                            Top =2580
                            Width =619
                            Height =195
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label12"
                            Caption ="MLID"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =1470
                            LayoutCachedTop =2580
                            LayoutCachedWidth =2089
                            LayoutCachedHeight =2775
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4545
                    Top =2580
                    Width =615
                    Height =195
                    ColumnOrder =8
                    FontSize =8
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="GroupID"
                    ControlSource ="GroupID"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =4545
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =2775
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =3930
                            Top =2580
                            Width =619
                            Height =195
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label13"
                            Caption ="GroupID"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =3930
                            LayoutCachedTop =2580
                            LayoutCachedWidth =4549
                            LayoutCachedHeight =2775
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5775
                    Top =2580
                    Width =720
                    Height =195
                    ColumnOrder =9
                    FontSize =8
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5775
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =2775
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =5160
                            Top =2580
                            Width =619
                            Height =195
                            FontSize =8
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label14"
                            Caption ="UnitID"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =5160
                            LayoutCachedTop =2580
                            LayoutCachedWidth =5779
                            LayoutCachedHeight =2775
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =1980
                    Width =1980
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="GetLink"
                    Caption ="Get Link Details"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x000000000000000090a4682190a468ae90a468e790a468e490a468a890a4681b ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a4682a90a468f390a468ff90a468d590a468d890a468ff90a468ed ,
                        0x90a4682100000000727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468cc90a468f390a4683f000000000000000090a4684890a468f6 ,
                        0x90a468bd00000000727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468ff90a468720000000000000000000000000000000090a46878 ,
                        0x90a468ff90a4682a727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468ff90a468ff90a468ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffff18 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff90a468ff727272ffffffffff727272ffffffffffffffffffffffff4b ,
                        0x90a4683690a468ff90a468960000000000000000000000000000000090a46890 ,
                        0x90a468ff90a468ff727272ffffffffff727272ffffffffffffffffffffffffa5 ,
                        0x0000000090a468ae90a468ff90a4689390a4682a90a4682a90a4689390a468ff ,
                        0x90a468b490a468ff727272ffffffffff727272fffffffffffffffffffffffffc ,
                        0xffffff4290a4681290a468c990a468ff90a468ff90a468ff90a468ff90a468c9 ,
                        0x90a4681590a468ff727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffedffffff5190a4680390a4685a90a4689c90a4689f90a4685d90a46803 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffab727272ff727272ff727272ff727272ff72727203 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff727272ffdcdcdcff777777f67272723f00000000 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff727272ff787878f47272723c0000000000000000 ,
                        0x0000000000000000727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff7272723c000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1980
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =2280
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8965045
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =1
                    PressedShade =65.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =33
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7140
                    Top =2520
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="DeleteLink"
                    Caption ="Delete Link"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7140
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2820
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =7
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =33
                    QuickStyleMask =-565
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =1980
                    Width =1980
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="AddLink"
                    Caption ="Add Another Link"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000a574451e ,
                        0xa57445b7a57445ffa5744596a5744524a57445a4a57445ffa57445ffa57445ff ,
                        0xa57445b7a574451e0000000000000000000000000000000000000000a57445ba ,
                        0xa57445f3a5744596a5744503a57445ffa57445ffa5744512a5744533a5744547 ,
                        0xa57445f3a57445c00000000000000000000000000000000000000000a57445f3 ,
                        0xa574459900000000a5744599a57445ff00000000a57445ffa574459900000000 ,
                        0xa5744599a57445f90000000000000000000000000000000000000000a57445ba ,
                        0xa57445f3a5744547a5744533a5744512a57445ffa57445ff00000000a574454c ,
                        0xa57445f3a57445c00000000000000000000000000000000000000000a574451b ,
                        0xa57445b1a57445ffa57445ffa57445ffa57445a4a5744524a5744596a57445ff ,
                        0xa57445b1a574451b000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1980
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2280
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8965045
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =1
                    PressedShade =65.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =33
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5460
                    Top =1980
                    Width =1980
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="SaveandClose"
                    Caption =" Save and Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ab598980ab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989800000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffd9b4caffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd9b4caffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5460
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =2280
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8965045
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =1
                    PressedShade =65.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =33
                    QuickStyleMask =-629
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "DocAddEditLinkForm_OLD.cls"
