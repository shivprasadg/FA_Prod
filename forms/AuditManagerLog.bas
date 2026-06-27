Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15758
    DatasheetFontHeight =11
    ItemSuffix =111
    Left =13500
    Top =12870
    Right =29565
    Bottom =26130
    RecSrcDt = Begin
        0xec532c4af2f8e540
    End
    RecordSource ="SELECT AssetID as [identity], ActionDate, [Action] as ActionDesc,Nz(OldLocation,"
        "[OldValue]) AS Prev, NZ(Newlocation,[NewValue]) AS New, SysUser AS [User], Field"
        "Label as FieldName, GroupID FROM vw_AuditUnits WHERE AssetID= 8842  ORDER BY Act"
        "ionDate DESC;"
    OnCurrent ="[Event Procedure]"
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
        Begin FormHeader
            Height =869
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Top =72
                    Width =15045
                    Height =270
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption =" ┌──── Search   ───────────────────────────────┬───                           ──"
                        "─────────────────┬─────────────────────────┬────────────────────────────────┬───"
                        "────────────────────┐"
                    GridlineColor =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =72
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =342
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =734
                    Width =12960
                    Height =135
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =734
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =869
                    TabIndex =11
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =45
                    Top =585
                    Width =1965
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label2"
                    Caption ="Action Date"
                    FontName ="Segoe UI"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =45
                    LayoutCachedTop =585
                    LayoutCachedWidth =2010
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    Left =2070
                    Top =585
                    Width =4665
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label3"
                    Caption ="   Action Taken "
                    FontName ="Segoe UI"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2070
                    LayoutCachedTop =585
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =8460
                    Top =585
                    Width =2880
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label4"
                    Caption ="From"
                    FontName ="Segoe UI"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8460
                    LayoutCachedTop =585
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =11400
                    Top =585
                    Width =2880
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label5"
                    Caption ="To"
                    FontName ="Segoe UI"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11400
                    LayoutCachedTop =585
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =14310
                    Top =585
                    Width =1440
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label6"
                    Caption ="User"
                    FontName ="Segoe UI"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14310
                    LayoutCachedTop =585
                    LayoutCachedWidth =15750
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =6795
                    Top =585
                    Width =1605
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label7"
                    Caption ="FieldName"
                    FontName ="Segoe UI"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6795
                    LayoutCachedTop =585
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2070
                    Top =345
                    Width =4665
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_ActionTaken"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2070
                    LayoutCachedTop =345
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =585
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8460
                    Top =345
                    Width =2880
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_Prev"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8460
                    LayoutCachedTop =345
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =585
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11400
                    Top =345
                    Width =2880
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_New"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11400
                    LayoutCachedTop =345
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =585
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =45
                    Top =345
                    Width =1965
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_ActionDate"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =45
                    LayoutCachedTop =345
                    LayoutCachedWidth =2010
                    LayoutCachedHeight =585
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14310
                    Top =345
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =9
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_SysUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SysUser FROM vw_AuditUnits WHERE AssetID= 8842 Group By SysUser ORDER BY "
                        "SysUser;"
                    OnClick ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14310
                    LayoutCachedTop =345
                    LayoutCachedWidth =15750
                    LayoutCachedHeight =585
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =3
                End
                Begin ComboBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6795
                    Top =345
                    Width =1605
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="Search_FieldName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FieldLabel as  FieldName FROM vw_AuditUnits WHERE AssetID= 8842 Group By "
                        "FieldLabel ORDER BY FieldLabel;"
                    OnClick ="[Event Procedure]"
                    Tag ="Search"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6795
                    LayoutCachedTop =345
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =585
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =11700
                    Width =288
                    Height =173
                    ColumnOrder =2
                    FontSize =9
                    Name ="GroupID"
                    ControlSource ="GroupID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =11700
                    LayoutCachedWidth =11988
                    LayoutCachedHeight =173
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =11988
                    Width =288
                    Height =173
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =1
                    Name ="ATFGroupSpecID"
                    ControlSource ="ATFGroupSpecID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =11988
                    LayoutCachedWidth =12276
                    LayoutCachedHeight =173
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =12276
                    Width =288
                    Height =173
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =2
                    Name ="SchID"
                    ControlSource ="SchID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12276
                    LayoutCachedWidth =12564
                    LayoutCachedHeight =173
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =215
                    Left =4380
                    Top =29
                    Width =1470
                    Height =233
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ClearResetForm"
                    Caption ="Clear Reset"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ab3255d65a0000000000000000000000000000000000000000 ,
                        0x3255d62d3255d693000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ae3255d6f93255d6360000000000000000000000003255d62d ,
                        0x3255d6db3255d61e000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6153255d6db3255d6f03255d630000000003255d6303255d6ea ,
                        0x3255d66300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d61b3255d6c33255d6ed3255d66f3255d6ea3255d6ae ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x727272ff727272ff000000003255d6033255d6c63255d6ff3255d6de3255d60c ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x7272727e000000003255d6153255d6ab3255d6ff3255d6cf3255d6bd3255d696 ,
                        0x3255d609000000000000000000000000000000000000000000000000727272ff ,
                        0x000000003255d64e3255d6ed3255d6ff3255d6b73255d60c000000003255d645 ,
                        0x3255d6a53255d6420000000000000000000000000000000000000000727272ff ,
                        0x000000003255d6753255d6de3255d65a00000000000000000000000000000000 ,
                        0x000000003255d6270000000000000000000000000000000000000000727272ff ,
                        0x7272728100000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000007272726c727272ff ,
                        0x727272ff727272ff727272780000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000007272724e727272fc727272ff ,
                        0x727272ff727272ff727272ff7272725a00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000072727236727272f6727272ff727272ff ,
                        0x727272ff727272ff727272ff727272f972727242000000000000000000000000 ,
                        0x00000000000000000000000072727224727272ea727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272f07272722d0000000000000000 ,
                        0x000000000000000000000000727272d2727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272e40000000000000000 ,
                        0x000000000000000000000000727272f0727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =29
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =262
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11385
                    Width =288
                    Height =173
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =10
                    Name ="SelectedID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =11385
                    LayoutCachedWidth =11673
                    LayoutCachedHeight =173
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                End
            End
        End
        Begin Section
            Height =322
            Name ="Detail"
            AlternateBackColor =16249583
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =1965
                    Height =270
                    FontSize =9
                    ForeColor =0
                    Name ="ActionDate"
                    ControlSource ="ActionDate"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2055
                    Top =30
                    Width =4665
                    Height =270
                    FontSize =9
                    TabIndex =1
                    ForeColor =0
                    Name ="ActionTaken"
                    ControlSource ="ActionDesc"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2055
                    LayoutCachedTop =30
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8445
                    Top =30
                    Width =2880
                    Height =270
                    FontSize =9
                    TabIndex =3
                    ForeColor =0
                    Name ="Prev"
                    ControlSource ="Prev"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8445
                    LayoutCachedTop =30
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11400
                    Top =30
                    Width =2880
                    Height =270
                    FontSize =9
                    TabIndex =4
                    ForeColor =0
                    Name ="New"
                    ControlSource ="New"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11400
                    LayoutCachedTop =30
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14310
                    Top =30
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="User"
                    ControlSource ="User"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =14310
                    LayoutCachedTop =30
                    LayoutCachedWidth =15750
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6780
                    Top =30
                    Width =1605
                    Height =270
                    FontSize =9
                    TabIndex =2
                    ForeColor =0
                    Name ="FieldName"
                    ControlSource ="FieldName"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =30
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Top =281
                    Width =12960
                    Height =36
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004600690065006c0064004e0061006d0065005d003d005b00740078007400 ,
                        0x430075007200720065006e0074005d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =281
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =317
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00180000005b00 ,
                        0x4600690065006c0064004e0061006d0065005d003d005b007400780074004300 ,
                        0x75007200720065006e0074005d00000000000000000000000000000000000000 ,
                        0x000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4320
                    Top =75
                    Width =765
                    Height =210
                    FontSize =8
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =0
                    Name ="FieldGroup"
                    ControlSource ="FieldGroup"
                    GridlineColor =0

                    LayoutCachedLeft =4320
                    LayoutCachedTop =75
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =285
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin FormFooter
            Height =510
            BackColor =14211288
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4740
                    Width =765
                    Height =210
                    FontSize =8
                    BorderColor =12632256
                    ForeColor =0
                    Name ="txtCurrent"
                    GridlineColor =0

                    LayoutCachedLeft =4740
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =210
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
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =100
                    Top =180
                    Width =839
                    Height =239
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BorderColor =1643706
                    ForeColor =5855577
                    Name ="Events"
                    FontName ="Segoe UI"
                    GridlineColor =0

                    LayoutCachedLeft =100
                    LayoutCachedTop =180
                    LayoutCachedWidth =939
                    LayoutCachedHeight =419
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =75
                            Width =900
                            Height =480
                            FontSize =8
                            BackColor =9211020
                            BorderColor =1643706
                            ForeColor =16777215
                            Name ="Label63"
                            Caption ="Events"
                            GridlineColor =0
                            LayoutCachedLeft =75
                            LayoutCachedWidth =975
                            LayoutCachedHeight =480
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6000
                    Top =105
                    Width =1215
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =10319446
                    Name ="ClearLabel"
                    Caption ="Reset Clear ALL"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0
                    LayoutCachedLeft =6000
                    LayoutCachedTop =105
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =330
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
    End
End
CodeBehindForm
' See "AuditManagerLog.cls"
