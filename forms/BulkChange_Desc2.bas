Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22579
    DatasheetFontHeight =11
    ItemSuffix =780
    Left =6660
    Top =5700
    Right =26205
    Bottom =9795
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xc98eb42e1528e640
    End
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            GridlineColor =12632256
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1155
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =960
                    Width =22579
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =960
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =1155
                    TabIndex =8
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =3
                    Left =60
                    Top =816
                    Width =864
                    Height =276
                    FontSize =9
                    Name ="Label73"
                    Caption ="♦AssetID"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =816
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2232
                    Top =816
                    Width =2112
                    Height =276
                    FontSize =9
                    Name ="Label74"
                    Caption ="VIN"
                    GroupTable =6
                    LayoutCachedLeft =2232
                    LayoutCachedTop =816
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =984
                    Top =816
                    Width =1188
                    Height =276
                    FontSize =9
                    Name ="Label75"
                    Caption ="Unit #"
                    GroupTable =6
                    LayoutCachedLeft =984
                    LayoutCachedTop =816
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
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
                    TextAlign =1
                    Left =5052
                    Top =816
                    Width =2880
                    Height =276
                    FontSize =9
                    Name ="Label83"
                    Caption ="Engine Type"
                    GroupTable =6
                    LayoutCachedLeft =5052
                    LayoutCachedTop =816
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5052
                    Top =504
                    Width =2880
                    Height =252
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =3
                    Name ="EngineCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_EngineType.EngineType, Lookup_EngineType.EngineDisplacement, Looku"
                        "p_EngineType.EngineHPRange FROM Lookup_EngineType;"
                    ColumnWidths ="1800;720;0"
                    StatusBarText ="lookup from Vendor table"
                    Tag ="Medit"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =5052
                    LayoutCachedTop =504
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5052
                    Top =180
                    Width =2880
                    Height =255
                    FontSize =8
                    FontWeight =700
                    Name ="EngineSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5052
                    LayoutCachedTop =180
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =435
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
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
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =16260
                    Top =300
                    Width =765
                    Height =210
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =7
                    Name ="txtCurrent"

                    LayoutCachedLeft =16260
                    LayoutCachedTop =300
                    LayoutCachedWidth =17025
                    LayoutCachedHeight =510
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17040
                    Top =285
                    Width =780
                    Height =450
                    ColumnOrder =7
                    FontSize =20
                    FontWeight =700
                    TabIndex =6
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =17040
                    LayoutCachedTop =285
                    LayoutCachedWidth =17820
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4404
                    Top =816
                    Width =576
                    Height =276
                    FontSize =9
                    Name ="Label189"
                    Caption ="Status"
                    GroupTable =6
                    LayoutCachedLeft =4404
                    LayoutCachedTop =816
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =4800
                    Height =540
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Spec | Engine | Trans | etc."
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =0
                End
                Begin EmptyCell
                    Left =60
                    Top =504
                    Width =864
                    Height =252
                    Name ="EmptyCell562"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =504
                    LayoutCachedWidth =924
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =984
                    Top =504
                    Width =1188
                    Height =252
                    Name ="EmptyCell563"
                    GroupTable =6
                    LayoutCachedLeft =984
                    LayoutCachedTop =504
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2232
                    Top =504
                    Width =2112
                    Height =252
                    Name ="EmptyCell564"
                    GroupTable =6
                    LayoutCachedLeft =2232
                    LayoutCachedTop =504
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =60
                    Top =180
                    Width =864
                    Height =255
                    Name ="EmptyCell568"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =924
                    LayoutCachedHeight =435
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =984
                    Top =180
                    Width =1188
                    Height =255
                    Name ="EmptyCell569"
                    GroupTable =6
                    LayoutCachedLeft =984
                    LayoutCachedTop =180
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =435
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2232
                    Top =180
                    Width =2112
                    Height =255
                    Name ="EmptyCell570"
                    GroupTable =6
                    LayoutCachedLeft =2232
                    LayoutCachedTop =180
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =435
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =4404
                    Top =180
                    Width =576
                    Height =255
                    Name ="EmptyCell657"
                    GroupTable =6
                    LayoutCachedLeft =4404
                    LayoutCachedTop =180
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =435
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =4404
                    Top =504
                    Width =576
                    Height =252
                    Name ="EmptyCell658"
                    GroupTable =6
                    LayoutCachedLeft =4404
                    LayoutCachedTop =504
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =15072
                    Top =180
                    Height =255
                    Name ="EmptyCell722"
                    GroupTable =6
                    LayoutCachedLeft =15072
                    LayoutCachedTop =180
                    LayoutCachedWidth =16512
                    LayoutCachedHeight =435
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =15072
                    Top =504
                    Height =252
                    Name ="EmptyCell723"
                    GroupTable =6
                    LayoutCachedLeft =15072
                    LayoutCachedTop =504
                    LayoutCachedWidth =16512
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12720
                    Top =504
                    Width =1296
                    Height =252
                    Name ="EmptyCell733"
                    GroupTable =6
                    LayoutCachedLeft =12720
                    LayoutCachedTop =504
                    LayoutCachedWidth =14016
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12720
                    Top =816
                    Width =1296
                    Height =276
                    Name ="EmptyCell734"
                    GroupTable =6
                    LayoutCachedLeft =12720
                    LayoutCachedTop =816
                    LayoutCachedWidth =14016
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12720
                    Top =180
                    Width =1296
                    Height =255
                    Name ="EmptyCell735"
                    GroupTable =6
                    LayoutCachedLeft =12720
                    LayoutCachedTop =180
                    LayoutCachedWidth =14016
                    LayoutCachedHeight =435
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =14076
                    Top =816
                    Width =936
                    Height =276
                    Name ="EmptyCell740"
                    GroupTable =6
                    LayoutCachedLeft =14076
                    LayoutCachedTop =816
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =15072
                    Top =816
                    Height =276
                    Name ="EmptyCell743"
                    GroupTable =6
                    LayoutCachedLeft =15072
                    LayoutCachedTop =816
                    LayoutCachedWidth =16512
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =14076
                    Top =504
                    Width =936
                    Height =252
                    Name ="EmptyCell744"
                    GroupTable =6
                    LayoutCachedLeft =14076
                    LayoutCachedTop =504
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =14076
                    Top =180
                    Width =936
                    Height =255
                    Name ="EmptyCell745"
                    GroupTable =6
                    LayoutCachedLeft =14076
                    LayoutCachedTop =180
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =435
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =9492
                    Top =816
                    Width =3168
                    Height =276
                    FontSize =9
                    Name ="Label752"
                    Caption ="Transmission Type"
                    GroupTable =6
                    LayoutCachedLeft =9492
                    LayoutCachedTop =816
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9492
                    Top =504
                    Width =3168
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="TransCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_Transmissions.TransmissionType, LookUp_Transmissions.TransmissionS"
                        "peeds, LookUp_Transmissions.TransmissionID FROM LookUp_Transmissions ORDER BY Lo"
                        "okUp_Transmissions.TransmissionType;"
                    ColumnWidths ="1800;720;0"
                    StatusBarText ="lookup from Vendor table"
                    Tag ="Medit"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =9492
                    LayoutCachedTop =504
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9492
                    Top =180
                    Width =3168
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="TransSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9492
                    LayoutCachedTop =180
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =435
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
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
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =8004
                    Top =816
                    Width =1416
                    Height =276
                    FontSize =9
                    Name ="Label764"
                    Caption ="Power Type"
                    GroupTable =6
                    LayoutCachedLeft =8004
                    LayoutCachedTop =816
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8004
                    Top =504
                    Width =1416
                    Height =252
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =15523798
                    Name ="PowerCbo"
                    ControlSource ="PowerType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_EnginePower.EngPowerType FROM Lookup_EnginePower;"
                    Tag ="Cloned"
                    GroupTable =6

                    LayoutCachedLeft =8004
                    LayoutCachedTop =504
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =756
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8004
                    Top =180
                    Width =1416
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="PowerSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8004
                    LayoutCachedTop =180
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =435
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
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
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =36
                    Width =864
                    Height =252
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =60
                    LayoutCachedTop =36
                    LayoutCachedWidth =924
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2232
                    Top =36
                    Width =2112
                    Height =252
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =6

                    LayoutCachedLeft =2232
                    LayoutCachedTop =36
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =984
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =6

                    LayoutCachedLeft =984
                    LayoutCachedTop =36
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5052
                    Top =36
                    Width =2880
                    Height =252
                    ColumnWidth =4290
                    FontSize =9
                    TabIndex =4
                    Name ="UnitEngineType"
                    ControlSource ="UnitEngineType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_EngineType.EngineType, Lookup_EngineType.EngineDisplacement, Looku"
                        "p_EngineType.EngineHPRange FROM Lookup_EngineType;"
                    ColumnWidths ="1800;720;0"
                    StatusBarText ="lookup from Vendor table"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =5052
                    LayoutCachedTop =36
                    LayoutCachedWidth =7932
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =302
                    Width =22579
                    Height =72
                    TabIndex =7
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4404
                    Top =36
                    Width =576
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    GroupTable =6

                    LayoutCachedLeft =4404
                    LayoutCachedTop =36
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12720
                    Top =36
                    Width =1296
                    Height =252
                    Name ="EmptyCell732"
                    GroupTable =6
                    LayoutCachedLeft =12720
                    LayoutCachedTop =36
                    LayoutCachedWidth =14016
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =14076
                    Top =36
                    Width =936
                    Height =252
                    Name ="EmptyCell741"
                    GroupTable =6
                    LayoutCachedLeft =14076
                    LayoutCachedTop =36
                    LayoutCachedWidth =15012
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =15072
                    Top =36
                    Height =252
                    Name ="EmptyCell742"
                    GroupTable =6
                    LayoutCachedLeft =15072
                    LayoutCachedTop =36
                    LayoutCachedWidth =16512
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9492
                    Top =36
                    Width =3168
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="UnitTransmissionType"
                    ControlSource ="UnitTransmissionType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_Transmissions.TransmissionType, LookUp_Transmissions.TransmissionS"
                        "peeds, LookUp_Transmissions.TransmissionID FROM LookUp_Transmissions ORDER BY Lo"
                        "okUp_Transmissions.TransmissionType;"
                    ColumnWidths ="1800;720;0"
                    StatusBarText ="lookup from Vendor table"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =9492
                    LayoutCachedTop =36
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8004
                    Top =36
                    Width =1416
                    Height =252
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BackColor =15523798
                    Name ="PowerType"
                    ControlSource ="PowerType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_EnginePower.EngPowerType FROM Lookup_EnginePower;"
                    Tag ="Cloned"
                    GroupTable =6

                    LayoutCachedLeft =8004
                    LayoutCachedTop =36
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =288
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
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
' See "BulkChange_Desc2.cls"
