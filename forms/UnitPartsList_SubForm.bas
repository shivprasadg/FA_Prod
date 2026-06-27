Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15408
    DatasheetFontHeight =11
    ItemSuffix =220
    Left =7935
    Top =9675
    Right =23070
    Bottom =15315
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xa227efcfcc21e540
    End
    RecordSource ="MakeModelManager_Parts"
    Caption ="UnitsIn_sub"
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
            Height =975
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =780
                    Width =15408
                    Height =195
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =780
                    LayoutCachedWidth =15408
                    LayoutCachedHeight =975
                    TabIndex =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5040
                    Top =660
                    Width =705
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5040
                    LayoutCachedTop =660
                    LayoutCachedWidth =5745
                    LayoutCachedHeight =900
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =13365
                    Top =660
                    Width =1740
                    Height =240
                    FontSize =9
                    Name ="ULoc_Label"
                    Caption ="Model"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13365
                    LayoutCachedTop =660
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =900
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5745
                    Top =660
                    Width =960
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitUnitNum_Label"
                    Caption ="Unit #"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5745
                    LayoutCachedTop =660
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =900
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =11895
                    Top =660
                    Width =1410
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitPlateNum_Label"
                    Caption ="Make"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11895
                    LayoutCachedTop =660
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =900
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9270
                    Top =660
                    Width =495
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitStatus_Label"
                    Caption ="Status"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9270
                    LayoutCachedTop =660
                    LayoutCachedWidth =9765
                    LayoutCachedHeight =900
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Top =660
                    Width =1650
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label86"
                    Caption ="Client"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =660
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =900
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =1650
                    Top =660
                    Width =1350
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label122"
                    Caption ="MLNo"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1650
                    LayoutCachedTop =660
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =900
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =3000
                    Top =660
                    Width =2040
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label123"
                    Caption ="Schedule | Group"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =660
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =900
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10320
                    Top =660
                    Width =1545
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label124"
                    Caption ="Sub Type"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10320
                    LayoutCachedTop =660
                    LayoutCachedWidth =11865
                    LayoutCachedHeight =900
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9765
                    Top =660
                    Width =555
                    Height =240
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label146"
                    Caption ="Year"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9765
                    LayoutCachedTop =660
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =900
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =13365
                    Top =60
                    Width =1740
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="ModelSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    BottomPadding =45
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13365
                    LayoutCachedTop =60
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =300
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
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =30
                    ListWidth =3240
                    Left =13365
                    Top =375
                    Width =1740
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =2
                    BackColor =15858167
                    BorderColor =12632256
                    Name ="ModelCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartModel.ChildPartModel, LookUp_ChildPartModel.ChildPartMake"
                        "ID, LookUp_ChildPartMake.ChildPartMake FROM LookUp_ChildPartModel INNER JOIN Loo"
                        "kUp_ChildPartMake ON LookUp_ChildPartModel.ChildPartMakeID = LookUp_ChildPartMak"
                        "e.ChildPartMakeID WHERE (((LookUp_ChildPartMake.ChildPartMake)=[Forms]![MakeMode"
                        "lManagerParts]![Make]));"
                    ColumnWidths ="2736;0;0"
                    Tag ="Medit"

                    LayoutCachedLeft =13365
                    LayoutCachedTop =375
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =630
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    Top =360
                    Width =1620
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =12632256
                    Name ="ClientCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName AS Client FROM MakeModelM"
                        "anagerParts_WorkTable INNER JOIN vw_SixKeys ON MakeModelManagerParts_WorkTable.U"
                        "nitID = vw_SixKeys.UnitID GROUP BY vw_SixKeys.clientID, vw_SixKeys.ClientShortNa"
                        "me ORDER BY vw_SixKeys.ClientShortName;"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"

                    LayoutCachedTop =360
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =6705
                    Top =660
                    Width =2565
                    Name ="EmptyCell205"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6705
                    LayoutCachedTop =660
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =900
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
            End
        End
        Begin Section
            Height =255
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5040
                    Width =705
                    Height =225
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =5745
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9270
                    Width =495
                    Height =225
                    ColumnWidth =840
                    ColumnOrder =7
                    FontSize =9
                    FontWeight =700
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9270
                    LayoutCachedWidth =9765
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5745
                    Width =960
                    Height =225
                    ColumnWidth =1500
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =5
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5745
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Width =2040
                    Height =225
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =3
                    Name ="SchGrp"
                    ControlSource ="SchGrp"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10320
                    Width =1545
                    Height =225
                    FontSize =9
                    TabIndex =8
                    BorderColor =-2147483615
                    Name ="UnitSubType"
                    ControlSource ="ChildType"
                    Format ="Short Date"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =11865
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1650
                    Width =1350
                    Height =225
                    FontSize =9
                    TabIndex =2
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1650
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Width =1650
                    Height =225
                    FontSize =9
                    TabIndex =1
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =1650
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9765
                    Width =555
                    Height =225
                    FontSize =9
                    TabIndex =7
                    Name ="Text140"
                    ControlSource ="ChildYear"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9765
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =3240
                    Left =13365
                    Width =1740
                    Height =225
                    FontSize =9
                    TabIndex =10
                    BackColor =15858167
                    BorderColor =12632256
                    Name ="ChildClass"
                    ControlSource ="ChildClass"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartModel.ChildPartModel, LookUp_ChildPartModel.ChildPartMake"
                        "ID, LookUp_ChildPartMake.ChildPartMake FROM LookUp_ChildPartModel INNER JOIN Loo"
                        "kUp_ChildPartMake ON LookUp_ChildPartModel.ChildPartMakeID = LookUp_ChildPartMak"
                        "e.ChildPartMakeID WHERE (((LookUp_ChildPartMake.ChildPartMake)=[Forms]![MakeMode"
                        "lManagerParts]![Make]));"
                    ColumnWidths ="2736;0;0"
                    Tag ="Medit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13365
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11895
                    Width =1410
                    Height =225
                    FontSize =9
                    TabIndex =9
                    BackColor =15858167
                    BorderColor =12632256
                    Name ="ChildMake"
                    ControlSource ="ChildMake"
                    Tag ="Medit"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11895
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6705
                    Width =2565
                    Height =225
                    FontSize =9
                    TabIndex =6
                    BackColor =15858167
                    BorderColor =12632256
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    Tag ="Medit"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6705
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
            End
        End
        Begin FormFooter
            Height =15
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =15
                    Height =0
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4155
                    LayoutCachedTop =15
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =15
                End
            End
        End
    End
End
CodeBehindForm
' See "UnitPartsList_SubForm.cls"
