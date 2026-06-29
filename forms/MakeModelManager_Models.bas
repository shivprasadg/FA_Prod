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
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10860
    DatasheetFontHeight =11
    ItemSuffix =229
    Left =20145
    Top =10215
    Right =-31651
    Bottom =16800
    ShortcutMenuBar ="SortClip"
    OrderBy ="[MakeModelManager_Models].[UnitModel]"
    RecSrcDt = Begin
        0xb3705865ced3e440
    End
    RecordSource ="SELECT LookUp_UnitModels.UnitMake, LookUp_UnitModels.UnitModel, LookUp_UnitModel"
        "s.MakeID, LookUp_UnitModels.INDX FROM LookUp_UnitModels ORDER BY LookUp_UnitMode"
        "ls.UnitModel;"
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
                    TextAlign =1
                    Left =2412
                    Top =180
                    Width =2280
                    Height =240
                    FontSize =10
                    Name ="Label147"
                    Caption ="Model"
                    GroupTable =9
                    LayoutCachedLeft =2412
                    LayoutCachedTop =180
                    LayoutCachedWidth =4692
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =180
                    Width =2160
                    Height =240
                    FontSize =10
                    Name ="Label3"
                    Caption ="Make"
                    GroupTable =9
                    LayoutCachedLeft =180
                    LayoutCachedTop =180
                    LayoutCachedWidth =2340
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
                    OverlapFlags =93
                    TextAlign =2
                    Left =4752
                    Top =180
                    Width =576
                    Height =240
                    FontSize =10
                    ForeColor =8872517
                    Name ="Label148"
                    Caption ="FA"
                    GroupTable =9
                    LayoutCachedLeft =4752
                    LayoutCachedTop =180
                    LayoutCachedWidth =5328
                    LayoutCachedHeight =420
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5388
                    Top =180
                    Width =576
                    Height =240
                    FontSize =10
                    ForeColor =8872517
                    Name ="Label149"
                    Caption ="Track"
                    GroupTable =9
                    LayoutCachedLeft =5388
                    LayoutCachedTop =180
                    LayoutCachedWidth =5964
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
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4725
                    Width =1185
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8872517
                    Name ="Label150"
                    Caption ="┌─ In Use─┐"
                    LayoutCachedLeft =4725
                    LayoutCachedWidth =5910
                    LayoutCachedHeight =240
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =6024
                    Top =180
                    Width =612
                    Name ="EmptyCell201"
                    GroupTable =9
                    LayoutCachedLeft =6024
                    LayoutCachedTop =180
                    LayoutCachedWidth =6636
                    LayoutCachedHeight =420
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =8856
                    Top =180
                    Width =1092
                    Name ="EmptyCell212"
                    GroupTable =9
                    LayoutCachedLeft =8856
                    LayoutCachedTop =180
                    LayoutCachedWidth =9948
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =8208
                    Top =180
                    Width =576
                    Name ="EmptyCell214"
                    GroupTable =9
                    LayoutCachedLeft =8208
                    LayoutCachedTop =180
                    LayoutCachedWidth =8784
                    LayoutCachedHeight =420
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =6708
                    Top =180
                    Name ="EmptyCell216"
                    GroupTable =9
                    LayoutCachedLeft =6708
                    LayoutCachedTop =180
                    LayoutCachedWidth =8148
                    LayoutCachedHeight =420
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =9
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =318
            Name ="Detail"
            AlternateBackColor =16249583
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =8856
                    Top =36
                    Width =1092
                    Height =252
                    FontSize =8
                    TabIndex =8
                    Name ="DeletedMe"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =8856
                    LayoutCachedTop =36
                    LayoutCachedWidth =9948
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
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =36
                    Width =2160
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    Name ="Make"
                    ControlSource ="=IIf(IsNull([INDX]),\"Type New Model Here►\",[UnitMake])"
                    GroupTable =9

                    LayoutCachedLeft =180
                    LayoutCachedTop =36
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
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
                    Left =6024
                    Top =36
                    Width =612
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="MakeID"
                    ControlSource ="MakeID"
                    GroupTable =9

                    LayoutCachedLeft =6024
                    LayoutCachedTop =36
                    LayoutCachedWidth =6636
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
                    Left =6708
                    Top =36
                    Height =252
                    FontSize =10
                    TabIndex =6
                    Name ="SaveUpdate"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =6708
                    LayoutCachedTop =36
                    LayoutCachedWidth =8148
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
                    Left =8208
                    Top =36
                    Width =576
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="INDX"
                    ControlSource ="INDX"
                    GroupTable =9

                    LayoutCachedLeft =8208
                    LayoutCachedTop =36
                    LayoutCachedWidth =8784
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
                    Left =2412
                    Top =36
                    Width =2280
                    Height =252
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    Name ="UnitModel"
                    ControlSource ="UnitModel"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GroupTable =9

                    LayoutCachedLeft =2412
                    LayoutCachedTop =36
                    LayoutCachedWidth =4692
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4752
                    Top =36
                    Width =576
                    Height =252
                    FontSize =9
                    TabIndex =2
                    ForeColor =8872517
                    Name ="Leased"
                    ControlSource ="=DLookUp(\"Leased\",\"Units_MakeModelsInUse\",\"UnitMake= '\" & [Make] & \"'  AN"
                        "D UnitModel= '\" & [UnitModel] & \"'\")"
                    Format ="Standard"
                    GroupTable =9

                    LayoutCachedLeft =4752
                    LayoutCachedTop =36
                    LayoutCachedWidth =5328
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5388
                    Top =36
                    Width =576
                    Height =252
                    FontSize =9
                    TabIndex =4
                    ForeColor =8872517
                    Name ="Tracked"
                    ControlSource ="=DLookUp(\"Tracked\",\"Units_MakeModelsInUse\",\"UnitMake= '\" & [Make] & \"'  A"
                        "ND UnitModel= '\" & [UnitModel] & \"'\")"
                    Format ="Standard"
                    GroupTable =9

                    LayoutCachedLeft =5388
                    LayoutCachedTop =36
                    LayoutCachedWidth =5964
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4752
                    Top =30
                    Width =432
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Text224"
                    ControlSource ="=IIf(IsNull([indx]),\"◄\",Null)"

                    LayoutCachedLeft =4752
                    LayoutCachedTop =30
                    LayoutCachedWidth =5184
                    LayoutCachedHeight =270
                    BackThemeColorIndex =1
                    ThemeFontIndex =1
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
' See "MakeModelManager_Models.cls"
