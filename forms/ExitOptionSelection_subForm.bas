Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4290
    DatasheetFontHeight =11
    ItemSuffix =132
    Left =16335
    Top =6435
    Right =21540
    Bottom =10725
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xa1ecb6ba1292e440
    End
    RecordSource ="Select * From UnitExitOption"
    DatasheetFontName ="Calibri"
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =945
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2190
                    Top =15
                    Width =1800
                    Height =255
                    FontSize =8
                    Name ="Label91"
                    Caption ="Term                 Months"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2190
                    LayoutCachedTop =15
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Top =30
                    Width =1395
                    Height =255
                    FontSize =10
                    BorderColor =12632256
                    Name ="OptionSelected"
                    ControlSource ="OptionSelected"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ExitOptionTypes.LeaseOptionType, ExitOptionTypes.LeaseOptionId FROM ExitO"
                        "ptionTypes;"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =30
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =285
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =630
                            Height =255
                            FontSize =8
                            Name ="Label226"
                            Caption ="Type"
                            FontName ="Segoe UI"
                            GroupTable =6
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =660
                            LayoutCachedHeight =285
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =345
                    Width =1395
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNotSent"
                    ControlSource ="NoticeSent"
                    Format ="Short Date"
                    StatusBarText ="Dt Notice sent to Assignee"
                    GroupTable =6

                    LayoutCachedLeft =720
                    LayoutCachedTop =345
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =345
                            Width =630
                            Height =240
                            FontSize =8
                            Name ="Label228"
                            Caption ="Sent"
                            FontName ="Segoe UI"
                            GroupTable =6
                            LayoutCachedLeft =30
                            LayoutCachedTop =345
                            LayoutCachedWidth =660
                            LayoutCachedHeight =585
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =645
                    Width =1395
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="Text256"
                    ControlSource ="NoticeConfrimed"
                    Format ="Short Date"
                    StatusBarText ="notice of Ex or Ext"
                    GroupTable =6

                    LayoutCachedLeft =720
                    LayoutCachedTop =645
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =915
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =645
                            Width =630
                            Height =270
                            FontSize =8
                            Name ="Label257"
                            Caption ="Replied"
                            FontName ="Segoe UI"
                            GroupTable =6
                            LayoutCachedLeft =30
                            LayoutCachedTop =645
                            LayoutCachedWidth =660
                            LayoutCachedHeight =915
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3765
                    Top =315
                    Width =255
                    TabIndex =4
                    Name ="Check70"
                    ControlSource ="ExchangePool"

                    LayoutCachedLeft =3765
                    LayoutCachedTop =315
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =555
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2190
                    Top =330
                    Width =1425
                    Height =240
                    FontSize =8
                    Name ="Label71"
                    Caption ="Matched to New?"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2190
                    LayoutCachedTop =330
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =570
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =4
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2610
                    Width =660
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =12632256
                    Name ="OptionTerm"
                    ControlSource ="OptionTerm"
                    RowSourceType ="Value List"
                    RowSource ="3;6;9;12;15;18;21;24;36;48;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23"
                        ";24;25;26;27;28;29;30;31;32;33;34;35;36"
                    Tag ="Medit"
                    GridlineColor =10921638

                    LayoutCachedLeft =2610
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =255
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =5760
                    Left =3210
                    Top =645
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =5
                    BorderColor =12632256
                    Name ="ExchangedByUnitId"
                    ControlSource ="ExchangedByUnitId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.UnitId, vw_SixKeys.Unitnum AS [Unit#], vw_SixKeys.MYear AS [Ye"
                        "ar], vw_SixKeys.Make, vw_SixKeys.Model, vw_SixKeys.Type FROM vw_SixKeys WHERE (("
                        "(vw_SixKeys.Status)=\"B\") AND ((vw_SixKeys.clientID)=[Forms]![UnitsForm]![clien"
                        "tID]) AND ((vw_SixKeys.MLNo)=[Forms]![UnitsForm]![MLNo]));"
                    ColumnWidths ="864;864;1008;1152;1008;1008"
                    StatusBarText ="notice of Ex or Ext"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3210
                    LayoutCachedTop =645
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =915
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2190
                            Top =645
                            Width =975
                            Height =270
                            FontSize =8
                            Name ="Label47"
                            Caption ="Replaced By"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2190
                            LayoutCachedTop =645
                            LayoutCachedWidth =3165
                            LayoutCachedHeight =915
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ExitOptionSelection_subForm.cls"
