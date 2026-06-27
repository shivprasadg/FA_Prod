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
    Width =11880
    DatasheetFontHeight =11
    ItemSuffix =144
    Left =7725
    Top =2715
    Right =20775
    Bottom =11340
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc0df42c1cdd3e440
    End
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
            FontSize =11
            FontName ="Calibri"
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
            Height =0
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =315
            Name ="Detail"
            AlternateBackColor =16249583
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Top =60
                    Width =1050
                    Height =255
                    FontSize =8
                    Name ="DeletedMe"
                    Caption ="Delete Me"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =315
                    RowStart =1
                    RowEnd =1
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
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =30
                    Width =3225
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    Name ="groupType"
                    ControlSource ="UnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitTypes.UnitType FROM LookUp_UnitTypes;"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =30
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1680
                    Top =30
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    Name ="UnitMake"
                    ControlSource ="UnitMake"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =30
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6225
                    Top =30
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    Name ="MakeID"
                    ControlSource ="MakeID"

                    LayoutCachedLeft =6225
                    LayoutCachedTop =30
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9780
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="SaveUpdate"
                    Caption ="Save Update"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9780
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedColor =8210719
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "MakeModelManager_Makes.cls"
