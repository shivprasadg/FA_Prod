Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12818
    DatasheetFontHeight =11
    ItemSuffix =240
    Left =5865
    Top =5325
    Right =20775
    Bottom =11520
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x248775785687e640
    End
    RecordSource ="SELECT LookUp_ChildPartModel.ChildPartModelID, LookUp_ChildPartModel.ChildPartMo"
        "del, LookUp_ChildPartModel.PartType, LookUp_ChildPartModel.ChildPartMakeID, Look"
        "Up_ChildPartMake.ChildPartMake, DCount(\"ChildPartID\",\"ChildParts\",\"ChildCla"
        "ss= '\" & [ChildPartModel] & \"' AND ChildMake= '\" & [ChildPartMake] & \"'\") A"
        "S InUse FROM LookUp_ChildPartModel INNER JOIN LookUp_ChildPartMake ON LookUp_Chi"
        "ldPartModel.ChildPartMakeID = LookUp_ChildPartMake.ChildPartMakeID ORDER BY Look"
        "Up_ChildPartModel.ChildPartModel;"
    Caption ="Add or Edit an Equipment Type"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =255
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
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
            Height =450
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =168
                    Top =180
                    Width =2592
                    Height =240
                    FontSize =10
                    Name ="Label3"
                    Caption ="                           Make"
                    GroupTable =9
                    LayoutCachedLeft =168
                    LayoutCachedTop =180
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =420
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7860
                    Top =180
                    Width =648
                    Height =240
                    FontSize =10
                    ForeColor =8872517
                    Name ="Label148"
                    Caption ="In Use"
                    GroupTable =9
                    LayoutCachedLeft =7860
                    LayoutCachedTop =180
                    LayoutCachedWidth =8508
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =8568
                    Top =180
                    Width =612
                    Name ="EmptyCell201"
                    GroupTable =9
                    LayoutCachedLeft =8568
                    LayoutCachedTop =180
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =420
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =11400
                    Top =180
                    Width =1092
                    Name ="EmptyCell212"
                    GroupTable =9
                    LayoutCachedLeft =11400
                    LayoutCachedTop =180
                    LayoutCachedWidth =12492
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =10752
                    Top =180
                    Width =576
                    Name ="EmptyCell214"
                    GroupTable =9
                    LayoutCachedLeft =10752
                    LayoutCachedTop =180
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =420
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =9252
                    Top =180
                    Name ="EmptyCell216"
                    GroupTable =9
                    LayoutCachedLeft =9252
                    LayoutCachedTop =180
                    LayoutCachedWidth =10692
                    LayoutCachedHeight =420
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5532
                    Top =180
                    Width =2268
                    FontSize =10
                    BorderColor =0
                    Name ="AddLabel1"
                    ControlSource ="=IIf(IsNull([ChildPartModelID]),\"▼Select A Type▼\",\"Part Type\")"
                    GroupTable =9

                    LayoutCachedLeft =5532
                    LayoutCachedTop =180
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =420
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2832
                            Top =180
                            Width =2628
                            Height =240
                            FontSize =10
                            Name ="Label147"
                            Caption ="Model"
                            GroupTable =9
                            LayoutCachedLeft =2832
                            LayoutCachedTop =180
                            LayoutCachedWidth =5460
                            LayoutCachedHeight =420
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =318
            Name ="Detail"
            AlternateBackColor =16249583
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5532
                    Top =36
                    Width =2268
                    Height =252
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    Name ="PartType"
                    ControlSource ="PartType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartType.ChildPartType FROM LookUp_ChildPartType;"
                    GroupTable =9

                    LayoutCachedLeft =5532
                    LayoutCachedTop =36
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11400
                    Top =36
                    Width =1092
                    Height =252
                    FontSize =8
                    TabIndex =8
                    Name ="DeletedMe"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =11400
                    LayoutCachedTop =36
                    LayoutCachedWidth =12492
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =1643706
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =9
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =168
                    Top =36
                    Width =2592
                    Height =252
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    Name ="UnitMake"
                    ControlSource ="=IIf(IsNull([ChildPartModelID]),\"Type New Model Here►\",[ChildPartmake])"
                    GroupTable =9

                    LayoutCachedLeft =168
                    LayoutCachedTop =36
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8568
                    Top =36
                    Width =612
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="MakeID"
                    ControlSource ="ChildPartMakeID"
                    GroupTable =9

                    LayoutCachedLeft =8568
                    LayoutCachedTop =36
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9252
                    Top =36
                    Height =252
                    FontSize =10
                    TabIndex =6
                    Name ="SaveUpdate"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =9252
                    LayoutCachedTop =36
                    LayoutCachedWidth =10692
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedColor =8210719
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =9
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10752
                    Top =36
                    Width =576
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="ChildPartModelID"
                    ControlSource ="ChildPartModelID"
                    GroupTable =9

                    LayoutCachedLeft =10752
                    LayoutCachedTop =36
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2832
                    Top =36
                    Width =2628
                    Height =252
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    Name ="ChildPartModel"
                    ControlSource ="ChildPartModel"
                    OnChange ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =2832
                    LayoutCachedTop =36
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7920
                    Top =14
                    Width =375
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8210719
                    Name ="AddLabel2"
                    ControlSource ="=IIf(IsNull([ChildPartModelID]),\"◄\",Null)"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =14
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =254
                    BackThemeColorIndex =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7860
                    Top =36
                    Width =648
                    Height =252
                    FontSize =9
                    TabIndex =4
                    ForeColor =8872517
                    Name ="Leased"
                    ControlSource ="InUse"
                    Format ="Standard"
                    GroupTable =9

                    LayoutCachedLeft =7860
                    LayoutCachedTop =36
                    LayoutCachedWidth =8508
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
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
' See "MakeModelManagerParts_Models.cls"
