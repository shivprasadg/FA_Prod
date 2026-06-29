Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20160
    DatasheetFontHeight =11
    ItemSuffix =816
    Left =4860
    Top =4200
    Right =26475
    Bottom =17010
    OnUnload ="[Event Procedure]"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xe0f9941dd97fe540
    End
    RecordSource ="SELECT UnitInspections.* FROM UnitInspections INNER JOIN UnitInsp_WorkTable ON U"
        "nitInspections.InspnID = UnitInsp_WorkTable.InspnID ORDER BY UnitInspections.Ass"
        "etID, UnitInspections.InspNum DESC;"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1266
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =901
                    Width =20100
                    Height =313
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =901
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =1214
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =600
                    Width =20100
                    Height =285
                    Name ="Box47"
                    LayoutCachedTop =600
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =885
                    BackThemeColorIndex =1
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =924
                    Top =948
                    Width =828
                    Height =288
                    FontSize =10
                    Name ="Label1"
                    Caption ="Unit #"
                    GroupTable =1
                    LayoutCachedLeft =924
                    LayoutCachedTop =948
                    LayoutCachedWidth =1752
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GroupTable =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =1812
                    Top =948
                    Width =2016
                    Height =288
                    FontSize =10
                    Name ="Label2"
                    Caption ="VIN"
                    GroupTable =1
                    LayoutCachedLeft =1812
                    LayoutCachedTop =948
                    LayoutCachedWidth =3828
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =36
                    Top =948
                    Width =0
                    Height =288
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =948
                    LayoutCachedWidth =36
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =924
                    Top =300
                    Width =828
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="send2Excel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =924
                    LayoutCachedTop =300
                    LayoutCachedWidth =1752
                    LayoutCachedHeight =540
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =36
                    Top =612
                    Width =0
                    Height =276
                    Name ="EmptyCell68"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =612
                    LayoutCachedWidth =36
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =612
                    Width =756
                    Height =276
                    Name ="EmptyCell69"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =612
                    LayoutCachedWidth =852
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =924
                    Top =612
                    Width =828
                    Height =276
                    Name ="EmptyCell70"
                    GroupTable =1
                    LayoutCachedLeft =924
                    LayoutCachedTop =612
                    LayoutCachedWidth =1752
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1812
                    Top =612
                    Width =2016
                    Height =276
                    Name ="EmptyCell71"
                    GroupTable =1
                    LayoutCachedLeft =1812
                    LayoutCachedTop =612
                    LayoutCachedWidth =3828
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =300
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =300
                    LayoutCachedWidth =36
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1812
                    Top =300
                    Width =2016
                    Name ="EmptyCell99"
                    GroupTable =1
                    LayoutCachedLeft =1812
                    LayoutCachedTop =300
                    LayoutCachedWidth =3828
                    LayoutCachedHeight =540
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3900
                    Top =948
                    Width =408
                    Height =288
                    FontSize =8
                    Name ="Label383"
                    Caption ="Num"
                    GroupTable =1
                    LayoutCachedLeft =3900
                    LayoutCachedTop =948
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7848
                    Top =948
                    Width =828
                    Height =288
                    FontSize =10
                    BackColor =9732411
                    Name ="Label426"
                    Caption ="Hours"
                    GroupTable =1
                    LayoutCachedLeft =7848
                    LayoutCachedTop =948
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12876
                    Top =300
                    Width =372
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="CopySUFmileage2"
                    Caption ="▼Copy SUF"
                    ControlTipText ="Copy the SUF Mileage to Inspected Mileage (only if missing data)"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12876
                    LayoutCachedTop =300
                    LayoutCachedWidth =13248
                    LayoutCachedHeight =540
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10040879
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10040879
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12204
                    Top =948
                    Width =600
                    Height =288
                    FontSize =10
                    ForeColor =6108695
                    Name ="Label511"
                    Caption ="Web"
                    GroupTable =1
                    LayoutCachedLeft =12204
                    LayoutCachedTop =948
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =13320
                    Top =612
                    Width =4560
                    Height =276
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="GetDekraData1"
                    Caption ="Get Insp1 Data"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13320
                    LayoutCachedTop =612
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26265
                    BorderColor =6108695
                    HoverColor =39372
                    PressedColor =26265
                    HoverForeColor =2646351
                    PressedForeColor =2646351
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =4380
                    Top =612
                    Width =1296
                    Height =276
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =8
                    BorderColor =9211020
                    Name ="iVendorCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorOb"
                        "solete FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.Vendo"
                        "rSNm;"
                    ColumnWidths ="0;10800;2880"
                    ControlTipText ="This value will be posted to the all of the units listed if you press the submit"
                        " button above"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =4380
                    LayoutCachedTop =612
                    LayoutCachedWidth =5676
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =300
                    Width =1296
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="iVendorIDSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =300
                    LayoutCachedWidth =5676
                    LayoutCachedHeight =540
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4380
                    Top =948
                    Width =1296
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label533"
                    Caption ="Vendor"
                    GroupTable =1
                    LayoutCachedLeft =4380
                    LayoutCachedTop =948
                    LayoutCachedWidth =5676
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5748
                    Top =948
                    Width =996
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label543"
                    Caption ="Insp Date "
                    GroupTable =1
                    LayoutCachedLeft =5748
                    LayoutCachedTop =948
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6816
                    Top =948
                    Width =960
                    Height =288
                    FontSize =10
                    BackColor =9732411
                    Name ="Label553"
                    Caption ="ODO"
                    GroupTable =1
                    LayoutCachedLeft =6816
                    LayoutCachedTop =948
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8748
                    Top =948
                    Width =1020
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label563"
                    Caption ="Result "
                    GroupTable =1
                    LayoutCachedLeft =8748
                    LayoutCachedTop =948
                    LayoutCachedWidth =9768
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8748
                    Top =300
                    Width =1020
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="iResultIDSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8748
                    LayoutCachedTop =300
                    LayoutCachedWidth =9768
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7848
                    Top =300
                    Width =828
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="GetDekraData2"
                    Caption ="Get Insp2 Data"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7848
                    LayoutCachedTop =300
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =26265
                    BorderColor =6108695
                    HoverColor =39372
                    PressedColor =26265
                    HoverForeColor =2646351
                    PressedForeColor =2646351
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-629
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9840
                    Top =612
                    Width =1080
                    Height =276
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =10
                    BorderColor =9211020
                    Name ="iPassedDateCBO"
                    Format ="Short Date"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1

                    LayoutCachedLeft =9840
                    LayoutCachedTop =612
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9840
                    Top =300
                    Width =1080
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="iPassedDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9840
                    LayoutCachedTop =300
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackShade =75.0
                    BorderColor =6108695
                    HoverThemeColorIndex =2
                    HoverShade =75.0
                    PressedThemeColorIndex =2
                    PressedShade =75.0
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =9840
                    Top =948
                    Width =1080
                    Height =288
                    FontSize =10
                    ForeColor =16777215
                    Name ="Label621"
                    Caption ="Passed"
                    GroupTable =1
                    LayoutCachedLeft =9840
                    LayoutCachedTop =948
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =8748
                    Top =612
                    Width =1020
                    Height =276
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =9
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000b5cb8800000000000200000003000000050000000101000000000000 ,
                        0xf7b5800000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="iResultIDCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionResults.InpsectionResultId, UnitInspectionResults.Inspectio"
                        "nResult FROM UnitInspectionResults;"
                    ColumnWidths ="0;1008"
                    GroupTable =1

                    LayoutCachedLeft =8748
                    LayoutCachedTop =612
                    LayoutCachedWidth =9768
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000b5cb8800010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010100 ,
                        0x0000000000f7b580000100000032000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =10992
                    Top =612
                    Width =1140
                    Height =276
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =11
                    Name ="iTypeIDCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionTypes.InspectionTypeId, UnitInspectionTypes.InspectionType "
                        "FROM UnitInspectionTypes;"
                    ColumnWidths ="0;1008"
                    GroupTable =1

                    LayoutCachedLeft =10992
                    LayoutCachedTop =612
                    LayoutCachedWidth =12132
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10992
                    Top =948
                    Width =1140
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label705"
                    Caption ="Type "
                    GroupTable =1
                    LayoutCachedLeft =10992
                    LayoutCachedTop =948
                    LayoutCachedWidth =12132
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3900
                    Top =612
                    Width =408
                    Height =276
                    Name ="EmptyCell718"
                    GroupTable =1
                    LayoutCachedLeft =3900
                    LayoutCachedTop =612
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6816
                    Top =612
                    Width =960
                    Height =276
                    Name ="EmptyCell721"
                    GroupTable =1
                    LayoutCachedLeft =6816
                    LayoutCachedTop =612
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7848
                    Top =612
                    Width =828
                    Height =276
                    Name ="EmptyCell732"
                    GroupTable =1
                    LayoutCachedLeft =7848
                    LayoutCachedTop =612
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3900
                    Top =300
                    Width =408
                    Name ="EmptyCell733"
                    GroupTable =1
                    LayoutCachedLeft =3900
                    LayoutCachedTop =300
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =540
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =96
                    Top =948
                    Width =756
                    Height =288
                    FontSize =10
                    Name ="Label734"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =948
                    LayoutCachedWidth =852
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =13320
                    Top =948
                    Width =4560
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label737"
                    Caption ="Altaas Note"
                    GroupTable =1
                    LayoutCachedLeft =13320
                    LayoutCachedTop =948
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6816
                    Top =300
                    Width =960
                    Name ="EmptyCell739"
                    GroupTable =1
                    LayoutCachedLeft =6816
                    LayoutCachedTop =300
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =540
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5748
                    Top =300
                    Width =996
                    Name ="EmptyCell744"
                    GroupTable =1
                    LayoutCachedLeft =5748
                    LayoutCachedTop =300
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =540
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10992
                    Top =300
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="iTypeIDSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10992
                    LayoutCachedTop =300
                    LayoutCachedWidth =12132
                    LayoutCachedHeight =540
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =10319446
                    BorderColor =6108695
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin EmptyCell
                    Left =5748
                    Top =612
                    Width =996
                    Height =276
                    Name ="EmptyCell751"
                    GroupTable =1
                    LayoutCachedLeft =5748
                    LayoutCachedTop =612
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =96
                    Top =300
                    Width =756
                    Name ="EmptyCell752"
                    GroupTable =1
                    LayoutCachedLeft =96
                    LayoutCachedTop =300
                    LayoutCachedWidth =852
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin EmptyCell
                    Left =12204
                    Top =612
                    Width =600
                    Height =276
                    Name ="EmptyCell773"
                    GroupTable =1
                    LayoutCachedLeft =12204
                    LayoutCachedTop =612
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12204
                    Top =300
                    Width =600
                    Name ="EmptyCell776"
                    GroupTable =1
                    LayoutCachedLeft =12204
                    LayoutCachedTop =300
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =540
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12876
                    Top =612
                    Width =372
                    Height =276
                    Name ="EmptyCell780"
                    GroupTable =1
                    LayoutCachedLeft =12876
                    LayoutCachedTop =612
                    LayoutCachedWidth =13248
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =13320
                    Top =300
                    Width =4560
                    Name ="EmptyCell789"
                    GroupTable =1
                    LayoutCachedLeft =13320
                    LayoutCachedTop =300
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =540
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19752
                    Top =300
                    Width =48
                    Name ="EmptyCell799"
                    GroupTable =1
                    LayoutCachedLeft =19752
                    LayoutCachedTop =300
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =540
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19752
                    Top =612
                    Width =48
                    Height =276
                    Name ="EmptyCell800"
                    GroupTable =1
                    LayoutCachedLeft =19752
                    LayoutCachedTop =612
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19752
                    Top =948
                    Width =48
                    Height =288
                    Name ="EmptyCell801"
                    GroupTable =1
                    LayoutCachedLeft =19752
                    LayoutCachedTop =948
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12876
                    Top =948
                    Width =372
                    Height =288
                    Name ="EmptyCell803"
                    GroupTable =1
                    LayoutCachedLeft =12876
                    LayoutCachedTop =948
                    LayoutCachedWidth =13248
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19308
                    Top =300
                    Width =372
                    Name ="EmptyCell804"
                    GroupTable =1
                    LayoutCachedLeft =19308
                    LayoutCachedTop =300
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =540
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19308
                    Top =612
                    Width =372
                    Height =276
                    Name ="EmptyCell805"
                    GroupTable =1
                    LayoutCachedLeft =19308
                    LayoutCachedTop =612
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =19308
                    Top =948
                    Width =372
                    Height =288
                    Name ="EmptyCell806"
                    GroupTable =1
                    LayoutCachedLeft =19308
                    LayoutCachedTop =948
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17940
                    Top =300
                    Width =1308
                    Name ="EmptyCell810"
                    GroupTable =1
                    LayoutCachedLeft =17940
                    LayoutCachedTop =300
                    LayoutCachedWidth =19248
                    LayoutCachedHeight =540
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17940
                    Top =612
                    Width =1308
                    Height =276
                    Name ="EmptyCell811"
                    GroupTable =1
                    LayoutCachedLeft =17940
                    LayoutCachedTop =612
                    LayoutCachedWidth =19248
                    LayoutCachedHeight =888
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =17940
                    Top =948
                    Width =1308
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    Name ="Label814"
                    Caption ="Requesterd By "
                    GroupTable =1
                    LayoutCachedLeft =17940
                    LayoutCachedTop =948
                    LayoutCachedWidth =19248
                    LayoutCachedHeight =1236
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =417
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =0
                    Height =276
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =36
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =924
                    Top =36
                    Width =828
                    Height =276
                    FontSize =9
                    TabIndex =2
                    Name ="UnitNum"
                    ControlSource ="iUnitNum"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =924
                    LayoutCachedTop =36
                    LayoutCachedWidth =1752
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =36
                    Width =2016
                    Height =276
                    FontSize =9
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="iVIN"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =1812
                    LayoutCachedTop =36
                    LayoutCachedWidth =3828
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =96
                    Top =36
                    Width =756
                    Height =276
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =10040879
                    Name ="UnitID"
                    ControlSource ="AssetId"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =96
                    LayoutCachedTop =36
                    LayoutCachedWidth =852
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7848
                    Top =36
                    Width =828
                    Height =276
                    FontSize =10
                    TabIndex =8
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="iHourRead"
                    ControlSource ="iHourRead"
                    Format ="Standard"
                    ControlTipText ="Unit's mileage when the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =7848
                    LayoutCachedTop =36
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3900
                    Top =36
                    Width =408
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =9211020
                    Name ="InspNum"
                    ControlSource ="InspNum"
                    ControlTipText ="Date the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =3900
                    LayoutCachedTop =36
                    LayoutCachedWidth =4308
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12204
                    Top =36
                    Width =600
                    Height =276
                    FontSize =10
                    TabIndex =12
                    BorderColor =10319446
                    Name ="InspectionLinkWeb"
                    ControlSource ="iWebLink"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =12204
                    LayoutCachedTop =36
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =4380
                    Top =36
                    Width =1296
                    Height =276
                    FontSize =10
                    TabIndex =5
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="iVendorID"
                    ControlSource ="iVendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorOb"
                        "solete FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.Vendo"
                        "rSNm;"
                    ColumnWidths ="0;10800;2880"
                    ControlTipText ="The Company that inspected the unit"
                    GroupTable =1
                    BottomPadding =105
                    AllowValueListEdits =0

                    LayoutCachedLeft =4380
                    LayoutCachedTop =36
                    LayoutCachedWidth =5676
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5748
                    Top =36
                    Width =996
                    Height =276
                    FontSize =9
                    TabIndex =6
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="iDate"
                    ControlSource ="iDate"
                    Format ="Short Date"
                    ControlTipText ="Date the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =5748
                    LayoutCachedTop =36
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6816
                    Top =36
                    Width =960
                    Height =276
                    FontSize =10
                    TabIndex =7
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="iOdoRead"
                    ControlSource ="iOdoRead"
                    Format ="Standard"
                    ControlTipText ="Unit's mileage when the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =6816
                    LayoutCachedTop =36
                    LayoutCachedWidth =7776
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9840
                    Top =36
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =10
                    BorderColor =9211020
                    Name ="iPassedDate"
                    ControlSource ="iPassedDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date the unit PASSED inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =9840
                    LayoutCachedTop =36
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =12876
                    Top =36
                    Width =372
                    Height =276
                    TabIndex =13
                    ForeColor =1643706
                    Name ="DeleteWebLink"
                    Caption ="Õ"
                    OnClick ="[Event Procedure]"
                    FontName ="Wingdings"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    BottomPadding =105
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =12876
                    LayoutCachedTop =36
                    LayoutCachedWidth =13248
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =1
                    HoverColor =-2
                    HoverShade =75.0
                    PressedColor =-2
                    HoverForeColor =1643706
                    PressedForeColor =1643706
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19752
                    Top =36
                    Width =48
                    Height =276
                    FontSize =10
                    TabIndex =17
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="InspnID"
                    ControlSource ="InspnID"
                    ControlTipText ="Unit's mileage when the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =19752
                    LayoutCachedTop =36
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =8748
                    Top =36
                    Width =1020
                    Height =276
                    FontSize =10
                    TabIndex =9
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000b5cb8800000000000200000003000000050000000101000000000000 ,
                        0xf7b5800000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="iResultID"
                    ControlSource ="iResultID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionResults.InpsectionResultId, UnitInspectionResults.Inspectio"
                        "nResult FROM UnitInspectionResults;"
                    ColumnWidths ="0;1008"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =8748
                    LayoutCachedTop =36
                    LayoutCachedWidth =9768
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000b5cb8800010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010100 ,
                        0x0000000000f7b580000100000032000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =10992
                    Top =36
                    Width =1140
                    Height =276
                    FontSize =10
                    TabIndex =11
                    Name ="iTypeID"
                    ControlSource ="iTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionTypes.InspectionTypeId, UnitInspectionTypes.InspectionType "
                        "FROM UnitInspectionTypes;"
                    ColumnWidths ="0;1008"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =10992
                    LayoutCachedTop =36
                    LayoutCachedWidth =12132
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =314
                    Width =20160
                    Height =59
                    FontSize =10
                    TabIndex =18
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =314
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =373
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13320
                    Top =36
                    Width =4560
                    Height =276
                    FontSize =9
                    TabIndex =14
                    BorderColor =9211020
                    Name ="iNotes"
                    ControlSource ="iNotes"
                    Tag ="Medit"
                    ControlTipText ="This Date will be posted to the all of the units listed if you press the submit "
                        "button above"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =13320
                    LayoutCachedTop =36
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =19308
                    Top =36
                    Width =372
                    Height =276
                    TabIndex =16
                    ForeColor =16777215
                    Name ="DeleteMe"
                    Caption ="DeleteMe"
                    OnClick ="[Event Procedure]"
                    FontName ="Wingdings"
                    ControlTipText ="Submit the Item below to all units listed"
                    GroupTable =1
                    BottomPadding =105
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d6943255d68300000000000000000000000000000000 ,
                        0x3255d6043255d681000000000000000000000000000000000000000000000000 ,
                        0x000000003255d60e3255d6f43255d6f63255d66a000000000000000000000000 ,
                        0x3255d6c13255d6b7000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d68d3255d6f83255d6f43255d652000000003255d6ba ,
                        0x3255d6ec3255d618000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000003255d6943255d6f63255d6f43255d6d23255d6f6 ,
                        0x3255d65900000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000003255d68d3255d6f83255d6f83255d6cd ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000000000003255d6753255d6f13255d6f83255d6eb3255d6f8 ,
                        0x3255d6c63255d614727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff423255d69b3255d6f83255d6f63255d6a1ffffff033255d65e ,
                        0x3255d6d93255d6ad727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affcead8da33255d6623255d6bf546acb52c9a582ccb48250ffc9a581cf ,
                        0xffffff3300000000727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb58454fecca988bbcdab8bb1ba8c5ffbb17d4affb17d4affb17d4aff ,
                        0xffffffff727272ff727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =19308
                    LayoutCachedTop =36
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    GroupTable =1
                    QuickStyle =31
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17940
                    Top =36
                    Width =1308
                    Height =276
                    FontSize =9
                    TabIndex =15
                    BackColor =16183539
                    BorderColor =9211020
                    Name ="Text808"
                    ControlSource ="RequestedBy"
                    ControlTipText ="Date the unit was inspected"
                    GroupTable =1
                    BottomPadding =105

                    LayoutCachedLeft =17940
                    LayoutCachedTop =36
                    LayoutCachedWidth =19248
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9540
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =9540
                    LayoutCachedWidth =9569
                    LayoutCachedHeight =270
                End
            End
        End
    End
End
CodeBehindForm
' See "ReMarketingMassMod_sub_Inspection2.cls"
