Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22579
    DatasheetFontHeight =11
    ItemSuffix =539
    Left =17670
    Top =1725
    Right =-25216
    Bottom =14430
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x28ceb32e1528e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            Height =1200
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =16560
                    Width =4380
                    Height =300
                    FontSize =8
                    Name ="Label538"
                    Caption ="┌───────────────── Title ───────────────────┐"
                    LayoutCachedLeft =16560
                    LayoutCachedWidth =20940
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =12165
                    Width =2565
                    Height =300
                    FontSize =8
                    Name ="Label126"
                    Caption ="┌───────── COA ─────────┐"
                    LayoutCachedLeft =12165
                    LayoutCachedWidth =14730
                    LayoutCachedHeight =300
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =1005
                    Width =22579
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =1005
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =1200
                    TabIndex =22
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =8783
                    Width =1455
                    Height =240
                    FontSize =8
                    Name ="Label521"
                    Caption ="┌ MSO | PLB Title ─┐"
                    LayoutCachedLeft =8783
                    LayoutCachedWidth =10238
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =30
                    Top =865
                    Width =990
                    Height =255
                    FontSize =9
                    Name ="Label73"
                    Caption ="♦AssetID"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =865
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2265
                    Top =865
                    Width =1875
                    Height =255
                    FontSize =9
                    Name ="Label74"
                    Caption ="VIN"
                    LayoutCachedLeft =2265
                    LayoutCachedTop =865
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =1080
                    Top =870
                    Width =1155
                    Height =255
                    FontSize =9
                    Name ="Label75"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =1080
                    LayoutCachedTop =870
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =1125
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11925
                    Top =865
                    Width =1008
                    Height =270
                    FontSize =9
                    Name ="Label82"
                    Caption ="COA Sent"
                    LayoutCachedLeft =11925
                    LayoutCachedTop =865
                    LayoutCachedWidth =12933
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16365
                    Top =865
                    Width =1200
                    Height =255
                    FontSize =9
                    Name ="Label84"
                    Caption ="Title App Sent"
                    LayoutCachedLeft =16365
                    LayoutCachedTop =865
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =15225
                    Top =865
                    Width =1008
                    Height =255
                    FontSize =9
                    Name ="Label103"
                    Caption ="POA Sent"
                    LayoutCachedLeft =15225
                    LayoutCachedTop =865
                    LayoutCachedWidth =16233
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =6000
                    Top =60
                    Width =765
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =60
                    LayoutCachedWidth =6765
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6000
                    Top =300
                    Width =780
                    Height =450
                    ColumnOrder =3
                    FontSize =20
                    FontWeight =700
                    TabIndex =7
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =300
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =750
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11925
                    Top =220
                    Width =1008
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="COASubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11925
                    LayoutCachedTop =220
                    LayoutCachedWidth =12933
                    LayoutCachedHeight =475
                    ColumnStart =9
                    ColumnEnd =9
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =11830108
                    BorderThemeColorIndex =1
                    HoverColor =3751056
                    PressedColor =0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11925
                    Top =550
                    Width =1008
                    Height =270
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =11
                    Name ="COASent"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =11925
                    LayoutCachedTop =550
                    LayoutCachedWidth =12933
                    LayoutCachedHeight =820
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15225
                    Top =220
                    Width =1008
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="POASubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15225
                    LayoutCachedTop =220
                    LayoutCachedWidth =16233
                    LayoutCachedHeight =460
                    ColumnStart =10
                    ColumnEnd =10
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15225
                    Top =550
                    Width =1008
                    Height =255
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =9
                    Name ="POASent"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =15225
                    LayoutCachedTop =550
                    LayoutCachedWidth =16233
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =16365
                    Top =220
                    Width =1200
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="TitleAppSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16365
                    LayoutCachedTop =220
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =460
                    ColumnStart =11
                    ColumnEnd =11
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16365
                    Top =550
                    Width =1200
                    Height =255
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =13
                    Name ="TitleAppSent"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =16365
                    LayoutCachedTop =550
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13020
                    Top =865
                    Width =1008
                    Height =270
                    FontSize =9
                    Name ="Label138"
                    Caption ="COA Copy"
                    LayoutCachedLeft =13020
                    LayoutCachedTop =865
                    LayoutCachedWidth =14028
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13050
                    Top =220
                    Width =930
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="COACopySubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13050
                    LayoutCachedTop =220
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =475
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =11830108
                    BorderThemeColorIndex =1
                    HoverColor =3751056
                    PressedColor =0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =550
                    Width =930
                    Height =270
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =12
                    Name ="COACopyCbo"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =13050
                    LayoutCachedTop =550
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =820
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =4
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8115
                    Top =865
                    Width =585
                    Height =255
                    FontSize =9
                    Name ="Label189"
                    Caption ="Status"
                    LayoutCachedLeft =8115
                    LayoutCachedTop =865
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8115
                    Top =220
                    Width =585
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="UnitStatusSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8115
                    LayoutCachedTop =220
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =460
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =8115
                    Top =550
                    Width =585
                    Height =255
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =8
                    Name ="UnitStatusCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="432;1440"
                    Tag ="Medit"

                    LayoutCachedLeft =8115
                    LayoutCachedTop =550
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Width =4305
                    Height =540
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="MSO|COA| POA|Title"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8820
                    Top =865
                    Width =1008
                    Height =270
                    FontSize =9
                    Name ="MSOCopyLbl"
                    Caption ="Copy"
                    RightPadding =38
                    LayoutCachedLeft =8820
                    LayoutCachedTop =865
                    LayoutCachedWidth =9828
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =8835
                    Top =220
                    Width =1008
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="UnitMSOSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8835
                    LayoutCachedTop =220
                    LayoutCachedWidth =9843
                    LayoutCachedHeight =460
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8820
                    Top =550
                    Width =1008
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =10
                    BackColor =15858167
                    Name ="UnitMSOCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =8820
                    LayoutCachedTop =550
                    LayoutCachedWidth =9828
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9855
                    Top =865
                    Width =1008
                    Height =270
                    FontSize =9
                    Name ="Label516"
                    Caption ="Original"
                    RightPadding =38
                    LayoutCachedLeft =9855
                    LayoutCachedTop =865
                    LayoutCachedWidth =10863
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =17640
                    Top =865
                    Width =2160
                    Height =255
                    FontSize =9
                    Name ="Label519"
                    Caption ="Title Number"
                    LayoutCachedLeft =17640
                    LayoutCachedTop =865
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4185
                    Top =865
                    Width =3825
                    Height =255
                    FontSize =9
                    Name ="Label520"
                    Caption ="Description"
                    LayoutCachedLeft =4185
                    LayoutCachedTop =865
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =1120
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin Line
                    OverlapFlags =95
                    Left =14467
                    Top =180
                    Width =0
                    Height =720
                    Name ="Line529"
                    LayoutCachedLeft =14467
                    LayoutCachedTop =180
                    LayoutCachedWidth =14467
                    LayoutCachedHeight =900
                End
                Begin Line
                    OverlapFlags =87
                    Left =10140
                    Top =120
                    Width =15
                    Height =720
                    Name ="Line530"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =120
                    LayoutCachedWidth =10155
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10980
                    Top =865
                    Width =765
                    Height =270
                    FontSize =9
                    Name ="Label178"
                    Caption ="is PLB"
                    RightPadding =38
                    LayoutCachedLeft =10980
                    LayoutCachedTop =865
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =432
                    Left =10980
                    Top =565
                    Width =765
                    Height =255
                    FontSize =9
                    TabIndex =14
                    BoundColumn =1
                    Name ="PLBcbo"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";-1;\"No\";0"
                    ColumnWidths ="432;0"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =10980
                    LayoutCachedTop =565
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =820
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =10980
                    Top =235
                    Width =765
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    Name ="PLBSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10980
                    LayoutCachedTop =235
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =475
                    ColumnStart =11
                    ColumnEnd =11
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =10980
                    Width =750
                    Height =240
                    FontSize =8
                    Name ="Label531"
                    Caption ="┌─ PLB ─┐"
                    LayoutCachedLeft =10980
                    LayoutCachedWidth =11730
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21360
                    Top =865
                    Width =1035
                    Height =240
                    FontSize =9
                    Name ="Label184"
                    Caption ="Accepted"
                    RightPadding =38
                    LayoutCachedLeft =21360
                    LayoutCachedTop =865
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =1105
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =21360
                    Top =220
                    Width =1035
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="AcceptedDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =21360
                    LayoutCachedTop =220
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =445
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21360
                    Top =550
                    Width =1035
                    Height =255
                    FontSize =9
                    TabIndex =17
                    Name ="AcceptedDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =21360
                    LayoutCachedTop =550
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14085
                    Top =220
                    Width =930
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    Name ="COAOrgSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14085
                    LayoutCachedTop =220
                    LayoutCachedWidth =15015
                    LayoutCachedHeight =475
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =2366701
                    BorderColor =16777215
                    HoverColor =3751056
                    PressedColor =0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14085
                    Top =550
                    Width =930
                    Height =270
                    FontSize =9
                    TabIndex =19
                    Name ="COAOrgCbo"
                    Format ="Short Date"
                    Tag ="Medit"

                    LayoutCachedLeft =14085
                    LayoutCachedTop =550
                    LayoutCachedWidth =15015
                    LayoutCachedHeight =820
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14085
                    Top =865
                    Width =930
                    Height =270
                    FontSize =8
                    FontWeight =700
                    ForeColor =2366701
                    Name ="COAOrgLabel"
                    Caption ="COA Orig"
                    LayoutCachedLeft =14085
                    LayoutCachedTop =865
                    LayoutCachedWidth =15015
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19860
                    Top =865
                    Width =1140
                    Height =270
                    FontSize =9
                    Name ="Label534"
                    Caption ="Lien Release"
                    RightPadding =38
                    LayoutCachedLeft =19860
                    LayoutCachedTop =865
                    LayoutCachedWidth =21000
                    LayoutCachedHeight =1135
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =19867
                    Top =220
                    Width =1140
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    Name ="LienReleaseSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19867
                    LayoutCachedTop =220
                    LayoutCachedWidth =21007
                    LayoutCachedHeight =445
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19867
                    Top =550
                    Width =1140
                    Height =255
                    FontSize =9
                    TabIndex =21
                    Name ="LienReleaseCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =19867
                    LayoutCachedTop =550
                    LayoutCachedWidth =21007
                    LayoutCachedHeight =805
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =30
                    Width =1008
                    Height =270
                    FontSize =9
                    TabIndex =8
                    BackColor =16381933
                    Name ="UnitCOACopy"
                    ControlSource ="UnitCofACopy"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    BottomPadding =38

                    LayoutCachedLeft =13050
                    LayoutCachedTop =30
                    LayoutCachedWidth =14058
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =990
                    Height =255
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =2280
                    Top =30
                    Width =1860
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    BottomPadding =38

                    LayoutCachedLeft =2280
                    LayoutCachedTop =30
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1095
                    Top =30
                    Width =1140
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1095
                    LayoutCachedTop =30
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11925
                    Top =30
                    Width =1008
                    Height =270
                    FontSize =9
                    TabIndex =7
                    BackColor =16381933
                    Name ="UnitCOASent"
                    ControlSource ="UnitCOASent"
                    Format ="Short Date"
                    BottomPadding =38

                    LayoutCachedLeft =11925
                    LayoutCachedTop =30
                    LayoutCachedWidth =12933
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16365
                    Top =30
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =10
                    Name ="UnitTASent"
                    ControlSource ="UnitTASent"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    BottomPadding =38

                    LayoutCachedLeft =16365
                    LayoutCachedTop =30
                    LayoutCachedWidth =17565
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15225
                    Top =30
                    Width =1008
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="UnitPOA"
                    ControlSource ="UnitPOA"
                    Format ="Short Date"
                    BottomPadding =38

                    LayoutCachedLeft =15225
                    LayoutCachedTop =30
                    LayoutCachedWidth =16233
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =10
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =302
                    Width =22579
                    Height =72
                    TabIndex =12
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =302
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =8115
                    Top =30
                    Width =585
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="720;1440"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    OnClick ="[Event Procedure]"
                    BottomPadding =38

                    LayoutCachedLeft =8115
                    LayoutCachedTop =30
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4200
                    Top =30
                    Width =3810
                    Height =255
                    FontSize =9
                    TabIndex =3
                    Name ="Text509"
                    ControlSource ="=DLookUp(\"DescriptionMed\",\"LookUp_UnitDesc\",\"AssetID= \" & [UnitID])"
                    StatusBarText ="VIN number must be unique in table"
                    BottomPadding =38

                    LayoutCachedLeft =4200
                    LayoutCachedTop =30
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14085
                    Top =30
                    Width =930
                    Height =270
                    FontSize =9
                    TabIndex =9
                    Name ="UnitCOABack"
                    ControlSource ="UnitCOABack"
                    Format ="Short Date"
                    BottomPadding =38

                    LayoutCachedLeft =14085
                    LayoutCachedTop =30
                    LayoutCachedWidth =15015
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =9
                    BackTint =20.0
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8820
                    Top =30
                    Width =1008
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BackColor =15858167
                    Name ="Text346"
                    ControlSource ="UnitMSOCopy"
                    Format ="Short Date"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8820
                    LayoutCachedTop =30
                    LayoutCachedWidth =9828
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9870
                    Top =30
                    Width =1008
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BackColor =15858167
                    Name ="UnitMSORecd"
                    ControlSource ="UnitMSORecd"
                    Format ="Short Date"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9870
                    LayoutCachedTop =30
                    LayoutCachedWidth =10878
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17640
                    Top =30
                    Width =2160
                    Height =255
                    FontSize =9
                    TabIndex =11
                    Name ="UnitTitleNumber"
                    ControlSource ="UnitTitleNumber"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =17640
                    LayoutCachedTop =30
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10980
                    Top =30
                    Width =765
                    Height =255
                    FontSize =9
                    TabIndex =14
                    Name ="UnitTaxExempt"
                    ControlSource ="UnitPLB"
                    Format ="Yes/No"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10980
                    LayoutCachedTop =30
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21360
                    Top =15
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =15
                    Name ="UnitAcceptDate"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =21360
                    LayoutCachedTop =15
                    LayoutCachedWidth =22395
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19860
                    Top =15
                    Width =1140
                    Height =270
                    FontSize =9
                    TabIndex =16
                    Name ="LienRelease"
                    ControlSource ="LienRelease"
                    Format ="Short Date"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =19860
                    LayoutCachedTop =15
                    LayoutCachedWidth =21000
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                End
            End
        End
        Begin FormFooter
            Height =60
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BulkChange_CoaMso.cls"
